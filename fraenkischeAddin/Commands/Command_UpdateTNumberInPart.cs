using System.Windows.Forms;
using Fraenkische.SWAddin.Services;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin.Commands
{
    public class Command_UpdateTNumberInPart : ICommand
    {
        private readonly SldWorks _swApp;

        private const string EXCEL_FILE_FILTER = "Excel Files|*.xlsx;*.xlsm;*.xls";

        public Command_UpdateTNumberInPart(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Load T-Number to PART";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                commandTitle: Title,
                tooltip: "Add T-Number to PART",
                iconI: 2, // např. 2. ikona ve tvém .bmp
                callback: Execute);
        }

        public void Execute()
        {
            var activeDoc = _swApp.IActiveDoc2 as ModelDoc2;
            if (activeDoc == null)
            {
                MessageBox.Show("This command only works on 'PART' documents.", "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //OpenFileDialog ofd = new OpenFileDialog
            //{
            //    Title = "Select 'TOOLBOX' Excel file",
            //    Filter = EXCEL_FILE_FILTER
            //};

            //if (ofd.ShowDialog() != DialogResult.OK) return;
            string excelPath = @"C:\Users\staffav\Fraenkische Rohrwerke Gebr. Kirchner GmbH & Co. KG\FIP_CZ_PEEN - Documents\Design Team\Toolshop_drawings.xlsm";

            var reader = new TNumberExcelReader(excelPath);
            var editor = new CustomPropertyEditor();
            var assigner = new TNumberAssigner(_swApp, reader, editor);

            assigner.UpdateTNumber(activeDoc);

        }
    }
}