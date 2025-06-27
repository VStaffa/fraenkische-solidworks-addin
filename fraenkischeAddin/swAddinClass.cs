using System;
using System.IO;
using System.Runtime.InteropServices;
using Fraenkische.SWAddin.Commands;
using Fraenkische.SWAddin.UI;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;

namespace Fraenkische.SWAddin
{

    [ComVisible(true)]
    //HOME GUID
    //[Guid("B59ACE60-12DE-4C96-9910-4A268557EF64")]

    //WORK GUID
    [Guid("E5F928C1-B502-41D2-BA19-D86E4AD34786")]

    public class SWAddinClass : SwAddin
    {
        private SldWorks swApp;
        private int swCookie;
        private TaskpaneView swTaskpaneView;
        private TaskpaneHostUI swTaskpaneHost;

        private CommandManagerService commandManager;
        private FeatureManager featureManager;

        public const string SWTASKPANE_PROGID = "fraenkischeAddin.Taskpane";

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (SldWorks)ThisSW;
            swCookie = Cookie;

            swApp.SetAddinCallbackInfo2(1, this, swCookie);

            //MessageBox.Show("Design Team Addin Connected Sucessfully!");

            commandManager = new CommandManagerService(swApp, swCookie);
            featureManager = new FeatureManager(swApp, commandManager);

            featureManager.RegisterFeatures();
            commandManager.Finalize();

            LoadUI();

            return true;
        }
        private void LoadUI()
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), "Resources", "AddinLogo.png");
            //MessageBox.Show(imagePath);

            swTaskpaneView = swApp.CreateTaskpaneView2(imagePath, "Smart Designer");
            swTaskpaneHost = (TaskpaneHostUI)swTaskpaneView.AddControl(SWAddinClass.SWTASKPANE_PROGID, string.Empty);

        }

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
            }

        }
    }
}
