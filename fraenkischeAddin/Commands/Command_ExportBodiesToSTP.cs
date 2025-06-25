using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    public class Command_ExportBodiesToSTP : ICommand
    {
        private readonly ISldWorks _swApp;

        public Command_ExportBodiesToSTP(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Export Colored Bodies";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                Title, 
                "Export each solid body as a colored STEP", 
                3, 
                Execute);
        }

        public void Execute()
        {

            var swModel = _swApp.ActiveDoc as ModelDoc2;
            var model = _swApp.IActiveDoc2 as PartDoc;


            if (model == null)
            {
                MessageBox.Show("This command only works on part documents.", "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object[] bodies = model.GetBodies2((int)swBodyType_e.swSolidBody, false);
            if (bodies == null || bodies.Length == 0)
            {
                MessageBox.Show("No solid bodies found in the part.","Warning.",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string targetFolder = ChooseFolder();
            if (string.IsNullOrWhiteSpace(targetFolder)) return;

            foreach (IBody2 body in bodies)
            {
                // Hide other bodies
                foreach (Body2 b in bodies) b.HideBody(true);
                body.HideBody(false);

                // Save as STEP
                string bodyName = body.Name.Replace("/", "_");
                string filePath = Path.Combine(targetFolder, bodyName + ".stp");
                //MessageBox.Show(filePath);

                int errors = 0, warnings = 0;

                swModel.SaveAs3(filePath,0,0);

            }

            foreach (Body2 body in bodies)
            {
                body.HideBody(false);
            }

            MessageBox.Show("Export completed.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
