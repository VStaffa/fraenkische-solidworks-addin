using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_1_BatchBOMtoExcelExport : ICommand
    {
        private readonly SldWorks _swApp;
        public CMD_1_BatchBOMtoExcelExport(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Batch Export BOMs";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                Title,
                "Open all drawings in a folder and export their BOMs to Excel",
                0,
                Execute);
        }

        public void Execute()
        {
            SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            string folderPath = PickFolder();
            if (string.IsNullOrEmpty(folderPath)) return;


            string[] files = Directory.GetFiles(folderPath, "*.SLDDRW");

            if (files.Length == 0)
            {
                MessageBox.Show("No drawing files found in the selected folder.", "No Files Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (string file in files)
            {
                ModelDoc2 swModel = swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocDRAWING,
                                                   (int)swOpenDocOptions_e.swOpenDocOptions_LoadLightweight,
                                                   "", 0, 0);

                if (swModel == null)
                {
                    MessageBox.Show($"Failed to open: {file}", "File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                ExportBOM(swApp, swModel);

                swApp.CloseDoc(swModel.GetTitle());
            }

            MessageBox.Show("Batch BOM export completed successfully.", "Process Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportBOM(SldWorks swApp, ModelDoc2 swModel)
        {
            Feature swFeat = swModel.FirstFeature();
            BomFeature swBomFeat = null;

            while (swFeat != null)
            {
                if (swFeat.GetTypeName2() == "BomFeat")
                {
                    swBomFeat = (BomFeature)swFeat.GetSpecificFeature2();
                    break;
                }

                swFeat = swFeat.GetNextFeature();
            }

            if (swBomFeat == null)
            {
                return;
            }

            object[] tableAnnotations = (object[])swBomFeat.GetTableAnnotations();

            string filePath = swModel.GetPathName();
            string dir = Path.GetDirectoryName(filePath);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string excelPath = Path.Combine(dir, nameWithoutExt + "_BOM.xls");

            foreach (object table in tableAnnotations)
            {
                IBomTableAnnotation ta = (IBomTableAnnotation)table;
                ta.SaveAsExcel(excelPath, false, false);
            }

        }

        private string PickFolder()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder containing drawings";
                DialogResult result = dlg.ShowDialog();
                return result == DialogResult.OK ? dlg.SelectedPath : string.Empty;
            }
        }

        private void DebugPrint(string msg)
        {
            System.Diagnostics.Debug.Print(msg);
            // Or log to file/messagebox if needed
        }
    }
}