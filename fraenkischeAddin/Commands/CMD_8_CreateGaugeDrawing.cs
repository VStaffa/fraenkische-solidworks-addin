using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic; // pro Interaction.InputBox
using Fraenkische.SWAddin.Services;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_8_CreateGaugeDrawing : ICommand
    {
        private readonly SldWorks _swApp;
        private readonly string _templatePath;

        public CMD_8_CreateGaugeDrawing(SldWorks swApp)
        {
            _swApp = swApp;

            // 1) Pevná cesta k šabloně v Resources\SWTemplates
            string addinFolder = Path.GetDirectoryName(typeof(CMD_8_CreateGaugeDrawing).Assembly.Location);
            _templatePath = Path.Combine(addinFolder, "Resources", "SWTemplates", "GaugeDrawTemp.DRWDOT");
        }

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                "Gauge Drawing",
                "Vytvoří výkres měrky s izometrickým pohledem a popisky",
                0,
                Execute
            );
        }

        public void Execute()
        {
            // 2) Kontrola aktivního dokumentu
            var model = _swApp.ActiveDoc as ModelDoc2;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                MessageBox.Show(
                    "Prosím otevřete sestavný dokument (Assembly).",
                    "Chybný dokument",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }
            var asm = (AssemblyDoc)model;

            // 3) Vstupy od uživatele
            string coNumber = Interaction.InputBox("CO-číslo:", "Vytvoření výkresu", "");
            if (string.IsNullOrWhiteSpace(coNumber)) return;

            string segCountStr = Interaction.InputBox("Počet segmentů (těl):", "Vytvoření výkresu", "0");
            if (!int.TryParse(segCountStr, out int segCount) || segCount < 0)
            {
                MessageBox.Show(
                    "Zadejte prosím platné číslo (0 nebo vyšší).",
                    "Neplatný počet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            // 4) Vytvoření výkresu z šablony
            DrawingDoc draw = (DrawingDoc)_swApp.NewDocument(
                _templatePath,
                (int)swDwgPaperSizes_e.swDwgPaperA3size,
                0, 0
            );
            // počkáme, až se dokument přepne
            ModelDoc2 drawModel = _swApp.ActiveDoc as ModelDoc2;
            if (drawModel == null)
            {
                MessageBox.Show("Nepodařilo se otevřít nový výkres.", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 5) Vložení izometrického pohledu
            string asmPath = model.GetPathName();

            SolidWorks.Interop.sldworks.View view;

            view = draw.CreateDrawViewFromModelView3(asmPath, "*Isometric", 0.2, 0.15, 0);

            draw.ForceRebuild();

            drawModel = (ModelDoc2)draw;

            // 6) Přidání popisků podle počtu segmentů
            double yPos = 0.28;
            for (int i = 0; i <= segCount; i++)
            {
                string noteText = $"{coNumber}_{i}";
                Note note = drawModel.InsertNote(noteText);
                var ann = note.GetAnnotation();
                ann.SetLeader3(true, 0, false, false, 0, 0);
                ann.SetPosition2(0.015, yPos, 0);
                ann.ApplyDefaultStyleAttributes();
                yPos -= 0.01;
            }


            string bomTPath = Path.Combine(
                Path.GetDirectoryName(typeof(CMD_8_CreateGaugeDrawing).Assembly.Location),
                "Resources",
                "SWTemplates",
                "GaugeBOMTemp.sldbomtbt"
            );

            view.InsertBomTable2(
                false,
                0.45,
                0.15,
                2,
                1,
                "",
                bomTPath
                );


            // 7) Hotovo – bez ukládání
            MessageBox.Show(
                $"Výkres pro „{coNumber}“ byl vytvořen (neuložen).",
                "Dokončeno",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}
