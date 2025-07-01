using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    internal class CommandManagerService
    {
        private readonly SldWorks _swApp;
        private readonly int _addinCookie;
        private readonly ICommandManager _cmdMgr;
        public List<Action> _callbacks = new List<Action>();
        private CommandGroup _cmdGroup;

        // MAIN COMMAND GROUP
        private const int MainCommandGroupId = 5;
        private const string MainTitle = "AutoKONSTRUKTÉR";
        private const string MainTooltip = "Seznam design funkci";

        public CommandManagerService(SldWorks swApp, int addinCookie)
        {
            _swApp = swApp;
            _addinCookie = addinCookie;
            _cmdMgr = _swApp.GetCommandManager(_addinCookie);
            CreateCommandGroup();
        }
        private void CreateCommandGroup()
        {
            int errors = 0;

            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                MainCommandGroupId,
                MainTitle,
                MainTooltip,
                "",
                -1,
                true,
                ref errors);

            if (_cmdGroup == null)
                throw new Exception("Failed to create command group.");
        }

        internal void AddCommand(string commandTitle, string tooltip, int iconI, Action callback)
        {
            int cmdId = _callbacks.Count; // 🔹 This assigns the command ID
            string callbackName = $"CallBackFunction({_callbacks.Count})";

            _callbacks.Add(callback);     // 🔹 Stores the callback at that index

            #region ICON SETUP
            // Přidej tlačítko do command group

            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Icons");
            
            string[] icons = new[]
            {
                Path.Combine(basePath, "Icons_20x20.bmp"),  // 20x20

            };

            string[] mainIcons = new[]
            {       

                Path.Combine(basePath, "mainIcons_32x32.bmp"), // 32x32
            };

            // set icons before AddCommandItem2
            _cmdGroup.IconList = icons;
            _cmdGroup.MainIconList = mainIcons;

            #endregion

            _cmdGroup.AddCommandItem2(
                commandTitle,
                0,
                tooltip,
                commandTitle,
                iconI,
                callbackName, // callback name
                "EnableCallback",
                cmdId,                   // this is the index that gets passed back
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
        }

        public void Finalize()
        {
            _cmdGroup.HasToolbar = true;
            _cmdGroup.ShowInDocumentType = (int)swDocTemplateTypes_e.swDocTemplateTypeNONE |
            (int)swDocTemplateTypes_e.swDocTemplateTypePART |
            (int)swDocTemplateTypes_e.swDocTemplateTypeASSEMBLY |
            (int)swDocTemplateTypes_e.swDocTemplateTypeDRAWING;
            ;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();
        }
        public int HandleCommandCall(int id)
        {
            try
            {
                _callbacks[id].Invoke();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in command: " + ex.Message);
            }
            return 0;
        }
        public int EnableCallback()
        {
            return 1; // 1 = povoleno
        }
        public void Dispose()
        {
            if (_cmdGroup != null)
            {
                _cmdMgr.RemoveCommandGroup(MainCommandGroupId);
                _cmdGroup = null;
            }
        }
    }
}