using System.IO;
using System.Windows.Forms;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    internal class TNumberScraper : ICommand
    {
        private readonly ISldWorks _swApp;

        public TNumberScraper(ISldWorks swApp)
        {
            _swApp = swApp;
        }
        public void Execute()
        {
            MessageBox.Show("TNumber Scraper - EXECUTE!");
        }

        public void Register(CommandManagerService cmdMgrService)
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(swAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), "Icon1.png");

            cmdMgrService.AddCommand(
                commandTitle: "Find T-Number",
                tooltip: "TNumber scraper tooltip",
                iconI: 3,
                callback: Execute);
        }
    }
}