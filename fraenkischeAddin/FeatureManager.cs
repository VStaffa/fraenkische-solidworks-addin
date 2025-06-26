using System.Collections.Generic;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    public class FeatureManager
    {
        private readonly SldWorks _swApp;
        private readonly CommandManagerService _cmdMgr;

        public FeatureManager(SldWorks swApp, CommandManagerService cmdMgr)
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
            new Command_UpdateTNumberInPart(_swApp),

            //Callback_1
            new Command_ExportBodiesToSTP(_swApp),

            //Callback_2
            new Command_LoadTNumberToBOM(_swApp),
            
            //Callback 3
            new Command_LoadTNumbersFromRobot(_swApp),
            // etc.
        };

            foreach (var feature in features)
                feature.Register(_cmdMgr);
        }
    }
}