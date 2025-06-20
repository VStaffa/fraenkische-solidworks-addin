﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Commands
{
    public class CommandManagerService
    {
        private readonly ISldWorks _swApp;
        private readonly int _addinCookie;
        private readonly ICommandManager _cmdMgr;
        private readonly List<Action> _callbacks = new List<Action>();
        private CommandGroup _cmdGroup;
        private int _currentCmdIndex = 0;

        // Libovolná identifikace skupiny příkazů
        private const int MainCommandGroupId = 5;
        private const string MainTitle = "SMART DESIGN COMMANDS";
        private const string MainTooltip = "Seznam funkci";
        public CommandManagerService(ISldWorks swApp, int addinCookie) 
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
            int cmdId = _callbacks.Count;
            _callbacks.Add(callback);

            // Přidej tlačítko do command group

            var basePath = Path.Combine(Path.GetDirectoryName(typeof(swAddinClass).Assembly.Location), "Resources");

            string[] icons = new[]
            {
                Path.Combine(basePath, "Icons_20x20.bmp"),  // 20x20
                Path.Combine(basePath, "Icons_32x32.bmp"), // 32x32
            };

            string[] mainIcons = new[]
{
                Path.Combine(basePath, "mainIcons_20x20.bmp"),  // 20x20
                Path.Combine(basePath, "mainIcons_32x32.bmp"), // 32x32
            };

            // set icons before AddCommandItem2
            _cmdGroup.IconList = icons;
            _cmdGroup.MainIconList = mainIcons;

            _cmdGroup.AddCommandItem2(
                commandTitle,
                0, // position
                tooltip,
                tooltip,
                iconI,
                nameof(HandleCommandCallback), // callback
                nameof(EnableCallback),        // enable callback
                cmdId,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
        }

        public void Finalize()
        {
            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();
        }
        public int HandleCommandCallback()
        {
            try
            {
                _callbacks[1].Invoke();
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser2(
                    "Chyba při vykonání příkazu: " + ex.Message,
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
            return 0;
        }
        public int EnableCallback()
        {
            return 0; // 1 = povoleno
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