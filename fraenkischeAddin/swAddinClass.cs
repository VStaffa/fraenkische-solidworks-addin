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
    [Guid("B59ACE60-12DE-4C96-9910-4A268557EF64")]

    public class swAddinClass : SwAddin
    {
        private ISldWorks swApp;
        private int swCookie;
        private TaskpaneView swTaskpaneView;
        private TaskpaneHostUI swTaskpaneHost;

        private CommandManagerService commandManager;
        private FeatureManager featureManager;

        public const string SWTASKPANE_PROGID = "fraenkischeAddin.Taskpane";

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (ISldWorks)ThisSW;
            swCookie = Cookie;

            var ok = swApp.SetAddinCallbackInfo2(0, this, swCookie);

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
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(swAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty),"Resources", "AddinLogo.png");
            //MessageBox.Show(imagePath);

            swTaskpaneView = swApp.CreateTaskpaneView2(imagePath, "Smart Designer");
            swTaskpaneHost = (TaskpaneHostUI)swTaskpaneView.AddControl(swAddinClass.SWTASKPANE_PROGID, string.Empty);

        }

        public bool DisconnectFromSW()
        {
            UnloadUI();
            commandManager.Dispose();
            return true;
        }

        private void UnloadUI()
        {
            swTaskpaneHost = null;
            swTaskpaneView.DeleteView();
            Marshal.ReleaseComObject(swTaskpaneView);
            swTaskpaneView = null;

        }
    }
}
