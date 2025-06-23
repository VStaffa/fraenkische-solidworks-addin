using System.Windows.Forms;
using Fraenkische.SWAddin.Commands;
using Fraenkische.SWAddin.Services;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin.Commands
{
    public class UpdateTNumbersCommand : ICommand
    {
        private readonly ISldWorks _swApp;

        public UpdateTNumbersCommand(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Update T-Numbers";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                commandTitle: Title,
                tooltip: "Update T-Numbers from Excel",
                iconI: 2, // např. 2. ikona ve tvém .bmp
                callback: Execute);
        }

        public void Execute()
        {
            var activeDoc = _swApp.IActiveDoc2 as ModelDoc2;
            if (activeDoc == null)
            {
                System.Windows.Forms.MessageBox.Show("OPEN A PART TO USE THIS FEATURE!","CHYBA!",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            string excelPath = @"C:\Users\staff\Desktop\excel.xlsx";

            var reader = new TNumberExcelReader(excelPath);
            var editor = new CustomPropertyEditor();
            var updater = new TNumberAssigner(_swApp, reader, editor);

            updater.UpdateTNumber(activeDoc);

            System.Windows.Forms.MessageBox.Show("T-Number update completed.");
        }
    }
}