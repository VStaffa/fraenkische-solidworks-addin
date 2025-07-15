using Fraenkische.SWAddin.UI;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_9_GenerateInfill : ICommand
    {
        private readonly SldWorks _swApp;

        public CMD_9_GenerateInfill(SldWorks swApp) => _swApp = swApp;

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                "Generovat výplň",
                "Vytvoří výplň dle výběru nebo ručně zadaných rozměrů",
                9,
                Execute);
        }

        public void Execute()
        {
            // 0) Do we have any document open?
            ModelDoc2 activeDoc = _swApp.ActiveDoc as ModelDoc2;
            AssemblyDoc swAssy = activeDoc as AssemblyDoc;
            if (activeDoc == null)
            {
                // no document → always manual mode
                using (var form = new GenerateInfillForm(_swApp))
                    form.ShowDialog();
                return;
            }

            // 1) If it’s an assembly and 4 valid faces are selected, go auto…
            if (activeDoc is AssemblyDoc && TryGetSelectedHoleSize(out int w, out int h, out Face2[] asmPair1, out Face2[] asmPair2))
            {
                // 1) vytvoříme formulář s předvyplněnými rozměry
                using (var form = new GenerateInfillForm(_swApp, w, h, asmPair1, asmPair2))
                {
                    form.ShowDialog();
                    if (!form.DialogResult.Equals(DialogResult.OK))
                        return;
                }
            }
            else
            {
                // 2) Otherwise manual
                using (var form = new GenerateInfillForm(_swApp))
                    form.ShowDialog();
            }
        }
        public bool TryGetSelectedHoleSize(
            out int widthMm,
            out int heightMm,
            out Face2[] asmPair1,
            out Face2[] asmPair2)
        {
            widthMm = heightMm = 0;
            asmPair1 = asmPair2 = null;

            var swModel = (ModelDoc2)_swApp.ActiveDoc;
            var selMgr = swModel.SelectionManager;
            if (selMgr.GetSelectedObjectCount2(-1) != 4) return false;

            // načteme Face2
            var faces = new List<Face2>();
            for (int i = 1; i <= 4; i++)
            {
                if (selMgr.GetSelectedObjectType3(i, -1) != (int)swSelectType_e.swSelFACES)
                    return false;
                var f = selMgr.GetSelectedObject6(i, -1) as Face2;
                if (f == null) return false;
                faces.Add(f);
            }

            // najdeme první pár rovnoběžných Face2
            int i1 = -1, j1 = -1;
            for (int i = 0; i < 4 && i1 < 0; i++)
            {
                for (int j = i + 1; j < 4; j++)
                {
                    if (AreParallel(faces[i], faces[j]))
                    {
                        i1 = i; j1 = j;
                        break;
                    }
                }
            }
            if (i1 < 0) return false;

            // zbývají dva pro druhý pár
            var rem = Enumerable.Range(0, 4).Where(k => k != i1 && k != j1).ToArray();
            int i2 = rem[0], j2 = rem[1];
            if (!AreParallel(faces[i2], faces[j2])) return false;

            // změříme vzdálenosti přes ClosestDistance
            object pA = null, pB = null, pC = null, pD = null;
            double d1 = swModel.ClosestDistance(faces[i1], faces[j1], out pA, out pB);
            double d2 = swModel.ClosestDistance(faces[i2], faces[j2], out pC, out pD);
            if (d1 <= 0 || d2 <= 0) return false;

            // na mm, šířka ≥ výška
            var dims = new[] { d1 * 1000, d2 * 1000 }
                .OrderByDescending(x => x)
                .Select(x => (int)Math.Round(x))
                .ToArray();
            widthMm = dims[0];
            heightMm = dims[1];

            asmPair1 = new[] { faces[i1], faces[j1] };
            asmPair2 = new[] { faces[i2], faces[j2] };
            return true;
        }
        /// <summary>
        /// Ov ěří, že dvě planární plochy mají rovnoběžné normály.
        /// </summary>
        private bool AreParallel(Face2 f1, Face2 f2)
        {
            // Získáme podkladovou Surface a její rovnicové parametry
            var surf1 = f1.GetSurface() as Surface;
            var surf2 = f2.GetSurface() as Surface;
            if (surf1 == null || surf2 == null)
                return false;

            // PlaneParams vrací COM pole: [A, B, C, D]
            var pars1 = surf1.PlaneParams as double[];
            var pars2 = surf2.PlaneParams as double[];
            if (pars1 == null || pars2 == null || pars1.Length < 3 || pars2.Length < 3)
                return false;

            // Normálové vektory obou rovin
            double ax = pars1[0], ay = pars1[1], az = pars1[2];
            double bx = pars2[0], by = pars2[1], bz = pars2[2];

            // Křížový součin n1 × n2
            double cx = ay * bz - az * by;
            double cy = az * bx - ax * bz;
            double cz = ax * by - ay * bx;

            // Plná rovnoběžnost => křížový součin je (0,0,0)
            const double tol = 1e-9;
            return Math.Abs(cx) < tol &&
                   Math.Abs(cy) < tol &&
                   Math.Abs(cz) < tol;
        }

        private class PlaneEq { public double A, B, C, D; }

        private PlaneEq GetPlaneEquation(Face2 face)
        {
            // předpoklad: Face2.GetSurface() vrací Surface s PlaneParams
            var surf = face.GetSurface() as Surface;
            // Surface.PlaneParams vrací pole 4 dvojek: [A,B,C,D]
            var pars = surf?.PlaneParams as double[];
            if (pars == null || pars.Length < 4)
                throw new InvalidCastException("Nelze získat parametry roviny.");
            return new PlaneEq { A = pars[0], B = pars[1], C = pars[2], D = pars[3] };
        }

        private double DistanceBetweenPlanes(PlaneEq p1, PlaneEq p2)
        {
            // 1) Délky původních normál
            double n1 = Math.Sqrt(p1.A * p1.A + p1.B * p1.B + p1.C * p1.C);
            double n2 = Math.Sqrt(p2.A * p2.A + p2.B * p2.B + p2.C * p2.C);
            if (n1 < 1e-9 || n2 < 1e-9)
                throw new InvalidOperationException("Neplatná rovina: nulová normála.");

            // 2) Normalizované rovnice: (Â, B̂, Ĉ, D̂) tak, že [Â,B̂,Ĉ] má délku 1
            double A1 = p1.A / n1, B1 = p1.B / n1, C1 = p1.C / n1, D1 = p1.D / n1;
            double A2 = p2.A / n2, B2 = p2.B / n2, C2 = p2.C / n2, D2 = p2.D / n2;

            // 3) Zarovnat směr druhé normály, aby se shodoval s první
            //    (pokud je skalární součin < 0, normála směřuje opačně)
            if (A1 * A2 + B1 * B2 + C1 * C2 < 0)
            {
                A2 = -A2; B2 = -B2; C2 = -C2; D2 = -D2;
            }

            // 4) Vzdálenost je pak |D2 - D1|
            return Math.Abs(D2 - D1);
        }

    }
}
