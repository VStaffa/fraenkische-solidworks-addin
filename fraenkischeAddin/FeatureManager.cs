using System.Collections.Generic;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    public class FeatureManager
    {
        private readonly ISldWorks _swApp;
        private readonly CommandManagerService _cmdMgr;

        public FeatureManager(ISldWorks swApp, CommandManagerService cmdMgr)
        {
            _swApp = swApp;
            _cmdMgr = cmdMgr;
        }

        public void RegisterFeatures()
        {
            List<ICommand> features = new List<ICommand>()
        {

                //SEM PRIDAVAT NOVE FUNKCE

                //Callback_0
                new TNumberScraper(_swApp),
                //Callback_1
                new GenerateInfill(_swApp),
                //Callback_2
                new TestCommand1(_swApp),
                //Callback 3
                new TestCommand2(_swApp),

            // new ExportToStepCommand(_swApp),
            // new RenameDocumentCommand(_swApp),
            // etc.
        };

            foreach (var feature in features)
                feature.Register(_cmdMgr);
        }
    }
}