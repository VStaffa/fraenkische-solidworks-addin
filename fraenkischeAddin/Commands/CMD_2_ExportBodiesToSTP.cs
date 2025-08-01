using Fraenkische.SWAddin.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Windows.Forms;
using System.Windows.Shapes;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_2_ExportBodiesToSTP : ICommand
    {
        private readonly SldWorks _swApp;

        public CMD_2_ExportBodiesToSTP(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Export Solid Bodies to STP";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                Title,
                "Export each solid body as STEP",
                1,
                Execute);
        }

        public void Execute()
        {
            var swModel = _swApp.ActiveDoc as ModelDoc2;
            PartDoc model = _swApp.IActiveDoc2 as PartDoc;

            if (model == null)
            {
                MessageBox.Show("This command only works on 'PART' documents.", "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetBarText.Write("Ready");
                return;
            }

            SetBarText.Write("Finding solid bodies...");
            object[] bodies = model.GetBodies2((int)swBodyType_e.swSolidBody, false);
            if (bodies == null || bodies.Length == 0)
            {
                MessageBox.Show("No solid bodies found in the part.", "Warning.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SetBarText.Write("Ready");
                return;
            }

            string targetFolder;

            targetFolder = Microsoft.VisualBasic.Interaction.InputBox("Zadejte cestu ke složce:",
            "Výběr složky");

            if (string.IsNullOrWhiteSpace(targetFolder))
            {
                MessageBox.Show("Zadání bylo zrušeno nebo je prázdné.");
                targetFolder = ChooseFolder();

                if (string.IsNullOrWhiteSpace(targetFolder))
                {
                    SetBarText.Write("Ready");
                    return;
                }
            }else
            {
                if (Directory.Exists(targetFolder))
                {
                    MessageBox.Show("Složka existuje:\n" + targetFolder);
                }
                else
                {
                    MessageBox.Show("Zadaná složka neexistuje:\n" + targetFolder);
                    targetFolder = ChooseFolder();


                    if (string.IsNullOrWhiteSpace(targetFolder))
                    {
                        SetBarText.Write("Ready");
                        return;
                    }
                }
            }



            int total = bodies.Length;
            int current = 0;

            foreach (IBody2 body in bodies)
            {
                current++;
                SetBarText.Write($"Exporting body {current} of {total}...");

                // Hide other bodies
                foreach (Body2 b in bodies) b.HideBody(true);
                body.HideBody(false);

                // Save as STEP
                string bodyName = body.Name.Replace("/", "_");
                string filePath = System.IO.Path.Combine(targetFolder, bodyName + ".stp");

                swModel.SaveAs3(filePath, 0, 0);
            }

            foreach (Body2 body in bodies)
            {
                body.HideBody(false);
            }

            SetBarText.Write("Export completed.");
            MessageBox.Show("Export completed.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            SetBarText.Write("Ready");
        }
        private string ChooseFolder()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            var dlg = folderBrowserDialog;
            dlg.Description = "Select export folder for STEP files";
            return dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
        }
    }
}
