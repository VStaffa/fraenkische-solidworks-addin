using System.IO;
using System.Windows.Forms;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    internal class TestCommand1 : ICommand
    {
        private readonly ISldWorks _swApp;

        public TestCommand1(ISldWorks swApp)
        {
            _swApp = swApp;
        }
        public void Execute()
        {
            MessageBox.Show("TestCommand 1 - EXECUTE!");
        }

        public void Register(CommandManagerService cmdMgrService)
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(swAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), "Icon1.png");

            cmdMgrService.AddCommand(
                commandTitle: "Test Command 1",
                tooltip: "Test 1 tooltip",
                iconI: 1,
                callback: Execute);
        }
    }
}