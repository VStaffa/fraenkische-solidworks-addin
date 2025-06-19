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

                new GenerateInfill(_swApp),
                new TNumberScraper(_swApp),
                new GenerateInfill(_swApp),
                new TNumberScraper(_swApp),
                new GenerateInfill(_swApp),
                new TNumberScraper(_swApp),
                new GenerateInfill(_swApp),
                new TNumberScraper(_swApp),
            // new ExportToStepCommand(_swApp),
            // new RenameDocumentCommand(_swApp),
            // etc.
        };

            foreach (var feature in features)
                feature.Register(_cmdMgr);
        }
    }
}