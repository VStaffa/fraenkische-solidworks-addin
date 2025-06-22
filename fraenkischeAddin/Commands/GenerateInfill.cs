using System.IO;
using System.Windows.Forms;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    internal class GenerateInfill : ICommand
    {
        private readonly ISldWorks _swApp;

        public GenerateInfill(ISldWorks swApp)
        {
            _swApp = swApp;
        }
        public void Execute()
        {
            MessageBox.Show("Generate Infill - EXECUTE!");
        }

        public void Register(CommandManagerService cmdMgrService)
        {

            cmdMgrService.AddCommand(
                commandTitle: "Generate Infill",
                tooltip: "Generate custom infill tooltip.",
                iconI: 0,
                callback: Execute);
        }
    }
}