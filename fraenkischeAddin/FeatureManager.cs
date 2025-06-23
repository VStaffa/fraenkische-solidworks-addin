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
            new UpdateTNumbersCommand(_swApp),
            //Callback_1
            //new UpdateTNumbersCommand(_swApp),
            //Callback_2
            //Callback 3
            // etc.
        };

            foreach (var feature in features)
                feature.Register(_cmdMgr);
        }
    }
}