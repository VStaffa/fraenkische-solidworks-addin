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
                tooltip: "Update missing T-Numbers from Excel",
                iconI: 2, // např. 2. ikona ve tvém .bmp
                callback: Execute);
        }

        public void Execute()
        {
            var activeDoc = _swApp.IActiveDoc2 as IAssemblyDoc;
            if (activeDoc == null)
            {
                System.Windows.Forms.MessageBox.Show("This command works only on assemblies.");
                return;
            }

            string excelPath = @"C:\Users\staff\Desktop\excel.xlsx";

            var reader = new TNumberExcelReader(excelPath);
            var editor = new CustomPropertyEditor();
            var updater = new AssemblyTNumberUpdater(_swApp, reader, editor);

            updater.UpdateAllComponentsTNumbers(activeDoc);

            System.Windows.Forms.MessageBox.Show("T-Number update completed.");
        }
    }
}