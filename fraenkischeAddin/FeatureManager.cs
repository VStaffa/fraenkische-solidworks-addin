using Fraenkische.SWAddin.Commands;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;

namespace Fraenkische.SWAddin
{
    internal class FeatureManager
    {
        private readonly SldWorks _swApp;
        private readonly CommandManagerService _cmdMgr;

        private readonly Dictionary<Type, ICommand> _featureMap = new Dictionary<Type, ICommand>();


        public FeatureManager(SldWorks swApp, CommandManagerService cmdMgr)
        {
            _swApp = swApp;
            _cmdMgr = cmdMgr;
        }
        public T Get<T>() where T : class, ICommand
        {
            return _featureMap.TryGetValue(typeof(T), out var cmd) ? cmd as T : null;
        }
        public void RegisterFeatures()
        {
            List<ICommand> features = new List<ICommand>()
        {

            //SEM PRIDAVAT NOVE FUNKCE

            //Callback_0 ++ 
            new CMD_1_BatchBOMtoExcelExport(_swApp),
            new CMD_2_ExportBodiesToSTP(_swApp),
            new CMD_3_LoadPriceFromRobot(_swApp),
            new CMD_4_LoadTNumbersFromRobot(_swApp),
            new CMD_5_MergeExcelFilesInFolder(),
            new CMD_6_UpdateLocalExcels(),
            new CMD_7_UpdateTNumberInPart(_swApp),
            new CMD_8_CreateGaugeDrawing(_swApp),
            new CMD_9_GenerateInfill(_swApp),

        };

            foreach (var feature in features)
            {
                _featureMap[feature.GetType()] = feature;
                feature.Register(_cmdMgr);
            }

        }
    }
}