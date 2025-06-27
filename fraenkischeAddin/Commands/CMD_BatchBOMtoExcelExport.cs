using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_BatchBOMtoExcelExport : ICommand
    {
        private readonly SldWorks _swApp;
        public CMD_BatchBOMtoExcelExport(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Batch Export BOMs";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                Title,
                "Open all drawings in a folder and export their BOMs to Excel",
                3,
                Execute);
        }

        public void Execute()
        {
            SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            string folderPath = PickFolder();
            if (string.IsNullOrEmpty(folderPath)) return;

            string[] files = Directory.GetFiles(folderPath, "*.SLDDRW");

            foreach (string file in files)
            {
                ModelDoc2 swModel = swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocDRAWING,
                                                   (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                                                   "",0, 0);

                if (swModel == null)
                {
                    DebugPrint($"Failed to open: {file}");
                    continue;
                }

                ExportBOM(swApp, swModel);

                swApp.CloseDoc(swModel.GetTitle());
            }
        }

        private void ExportBOM(SldWorks swApp, ModelDoc2 swModel)
        {
            Feature swFeat = swModel.FirstFeature();
            BomFeature swBomFeat = null;

            while (swFeat != null)
            {
                if (swFeat.Name == "Bill of Materials1")
                {
                    swBomFeat = (BomFeature)swFeat.GetSpecificFeature2();
                    DebugPrint($"Found BOM: {swFeat.Name}");
                    break;
                }

                swFeat = swFeat.GetNextFeature();
            }

            if (swBomFeat == null)
            {
                DebugPrint($"No BOM found in: {swModel.GetTitle()}");
                return;
            }

            object[] tableAnnotations = (object[])swBomFeat.GetTableAnnotations();

            string excelPath = Path.ChangeExtension(swModel.GetPathName(), "_BOM.xls");

            foreach (object table in tableAnnotations)
            {
                IBomTableAnnotation ta = (IBomTableAnnotation)table;
                ta.SaveAsExcel(excelPath, false, true);
            }

            DebugPrint($"Exported to: {excelPath}");
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
