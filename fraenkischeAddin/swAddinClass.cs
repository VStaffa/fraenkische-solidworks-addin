using Fraenkische.SWAddin.Commands;
using Fraenkische.SWAddin.UI;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Fraenkische.SWAddin
{

    [ComVisible(true)]
    [Guid("E5F928C1-B502-41D2-BA19-D86E4AD34786")]

    public class SWAddinClass : SwAddin
    {
        private SldWorks swApp;
        private int swCookie;
        private TaskpaneView swTaskpaneView;

        private CommandManagerService commandManager;
        private FeatureManager featureManager;
        private TaskpaneHostUI swTaskpaneHost;
        internal static Frame myFrame { get; private set; }

        public const string SWTASKPANE_PROGID = "fraenkischeAddin.Taskpane";

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (SldWorks)ThisSW;
            swCookie = Cookie;
            swApp.SetAddinCallbackInfo2(1, this, swCookie);

            //ADD COMMAND GROUP AND POPULATE COMMANDS
            commandManager = new CommandManagerService(swApp, swCookie);
            featureManager = new FeatureManager(swApp, commandManager);

            featureManager.RegisterFeatures();
            commandManager.Finalize();

            myFrame = swApp.Frame();

            //CREATE TASKPANE
            LoadUI();

            swApp.ActiveDocChangeNotify += OnActiveDocChanged;
            UpdateActiveDocumentName();
            return true;
        }

        #region CURRENT OPEN DOCUMENT HANDLING
        private int OnActiveDocChanged()
        {
            UpdateActiveDocumentName();
            return 0;
        }

        private void UpdateActiveDocumentName()
        {
            var doc = swApp.IActiveDoc2;
            string name = doc != null ? Path.GetFileName(doc.GetPathName()) : "(none)";
            swTaskpaneHost?.UpdateDocumentName(name);
        }
        #endregion

        #region TASKPANE BUTTON CLICK HANDLER

        private void LoadUI()
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), @"Resources\Icons\mainIcons_20x20.bmp");
            swTaskpaneView = swApp.CreateTaskpaneView2(imagePath, "Smart Designer");
            swTaskpaneHost = (TaskpaneHostUI)swTaskpaneView.AddControl(SWAddinClass.SWTASKPANE_PROGID, string.Empty);

            swTaskpaneHost.Width = 400;
            swTaskpaneHost.MinimumSize = new System.Drawing.Size (400, 0);

            #region MATCH TASKPANE UI TO COMMANDS
            swTaskpaneHost.cmd_1_Clicked += () =>
            {
                featureManager.Get<CMD_1_BatchBOMtoExcelExport>()?.Execute();
            };
            swTaskpaneHost.cmd_2_Clicked += () =>
            {
                featureManager.Get<CMD_2_ExportBodiesToSTP>()?.Execute();
            };
            swTaskpaneHost.cmd_3_Clicked += () =>
            {
                featureManager.Get<CMD_3_LoadPriceFromRobot>()?.Execute();
            };
            swTaskpaneHost.cmd_4_Clicked += () =>
            {
                featureManager.Get<CMD_4_LoadTNumbersFromRobot>()?.Execute();
            };
            swTaskpaneHost.cmd_5_Clicked += () =>
            {
                featureManager.Get<CMD_5_MergeExcelFilesInFolder>()?.Execute();
            };
            swTaskpaneHost.cmd_6_Clicked += () =>
            {
                featureManager.Get<CMD_6_CopyExcelsToDesktop>()?.Execute();
            };
            swTaskpaneHost.cmd_7_Clicked += () =>
            {
                featureManager.Get<CMD_7_UpdateTNumberInPart>()?.Execute();
            };
            swTaskpaneHost.cmd_8_Clicked += () =>
            {
                featureManager.Get<CMD_8_CreateGaugeDrawing>()?.Execute();
            };
            swTaskpaneHost.cmd_9_Clicked += () =>
            {
                featureManager.Get<CMD_9_GenerateInfill>()?.Execute();
            };  
            #endregion
        }

        #endregion

        #region Dsconnect from SolidWorks and UNLOAD
        public bool DisconnectFromSW()
        {
            UnloadUI();
            commandManager.Dispose();
            swApp = null;
            return true;
        }

        private void UnloadUI()
        {
            if (swTaskpaneView != null)
            {
                swTaskpaneHost = null;
                swTaskpaneView.DeleteView();
                Marshal.ReleaseComObject(swTaskpaneView);
                swTaskpaneView = null;
            }

        }
        #endregion

        #region ICOMMAND CALLBACK HANDLING

        // This method is called by SolidWorks when a command is executed
        //CALLBACK FOR EACH FEATURE
        public void CallBackFunction(string data)
        {
            int commandIndex = int.Parse(data);
            switch (commandIndex)
            {
                case 0:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 1:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 2:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 3:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 4:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 5:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 6:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 7:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 8:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 9:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 10:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                case 11:
                    commandManager.HandleCommandCall(commandIndex);
                    break;
                    
            }

        }
        #endregion
    }
}
