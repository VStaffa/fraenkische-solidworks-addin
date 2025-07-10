using Fraenkische.SWAddin.Core;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_9_GenerateInfill : ICommand
    {
        private readonly SldWorks _swApp;
        public CMD_9_GenerateInfill(SldWorks swApp) => _swApp = swApp;

        public void Register(CommandManagerService cmdMgr)
        {
            // dej sem jedinečné ID a ikonu podle vašeho resource
            cmdMgr.AddCommand(
              "Generovat výplň",
              "Otevře formulář pro generování výplně",
              8,
              Execute);
        }

        public void Execute()
        {
            var form = new UI.GenerateInfillForm(_swApp);
            form.ShowDialog();
        }
    }
}