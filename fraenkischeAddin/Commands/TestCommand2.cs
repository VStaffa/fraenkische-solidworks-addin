using System.IO;
using System.Windows.Forms;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    internal class TestCommand2 : ICommand
    {
        private readonly ISldWorks _swApp;

        public TestCommand2(ISldWorks swApp)
        {
            _swApp = swApp;
        }
        public void Execute()
        {
            MessageBox.Show("TestCommand 2 - EXECUTE!");
        }

        public void Register(CommandManagerService cmdMgrService)
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(swAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), "Icon1.png");

            cmdMgrService.AddCommand(
                commandTitle: "Test Command 2",
                tooltip: "Test 2 tooltip",
                iconI: 2,
                callback: Execute);
        }
    }
}