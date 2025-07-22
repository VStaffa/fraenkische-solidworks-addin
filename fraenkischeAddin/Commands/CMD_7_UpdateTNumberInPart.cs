using Fraenkische.SWAddin.Core;
using Fraenkische.SWAddin.Services;
using SolidWorks.Interop.sldworks;
using System.Windows.Forms;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_7_UpdateTNumberInPart : ICommand
    {
        private readonly SldWorks _swApp;

        private const string EXCEL_FILE_FILTER = "Excel Files|*.xlsx;*.xlsm;*.xls";

        public CMD_7_UpdateTNumberInPart(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Load T-Number to PART";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                commandTitle: Title,
                tooltip: "Add T-Number to PART",
                iconI: 6, // např. 2. ikona ve tvém .bmp
                callback: Execute);
        }
        public void Execute()
        {
            ModelDoc2 activeDoc = _swApp.IActiveDoc2;

            if (activeDoc == null)
            {
                MessageBox.Show("Open a document", "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetBarText.Write("Ready");
                return;
            }

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = EXCEL_FILE_FILTER;
                openFileDialog.Title = "Select 'TOOLBOX' Excel File";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    SetBarText.Write("Ready");
                    return;
                }

                string excelPath = openFileDialog.FileName;

                SetBarText.Write("Reading T-Number from Excel...");
                var reader = new TNumberExcelReader(excelPath);
                var editor = new CustomPropertyEditor();
                var assigner = new TNumberAssigner(_swApp, reader, editor);

                SetBarText.Write("Assigning T-Number to part...");
                assigner.UpdateTNumber(activeDoc);

                SetBarText.Write("Ready");
            }
        }
    }
}