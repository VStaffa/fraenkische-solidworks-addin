using System.Collections.Generic;
using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin
{
    internal class FeatureManager
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
            new CMD_UpdateTNumberInPart(_swApp),

            //Callback_1
            new CMD_ExportBodiesToSTP(_swApp),

            //Callback_2
            //new Command_LoadTNumberToBOM(_swApp),
            
            //Callback 3
            new CMD_LoadTNumbersFromRobot(_swApp),
          
            //Callback 4
            new CMD_BatchBOMtoExcelExport(_swApp),

            //Callback 5
            new CMD_MergeExcelFilesInFolder(),

            //Callback 5
            new CMD_LoadPriceFromRobot(_swApp),

            //Callback 6
            new CMD_UpdateLocalExcels(),
            // etc.

        };

            foreach (var feature in features)
                feature.Register(_cmdMgr);
        }
    }
}