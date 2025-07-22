using System;
using System.IO;
using System.Net;
using System.Windows;
using Fraenkische.SWAddin.Core;


namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_6_CopyExcelsToDesktop : ICommand
    {
        public void Register(CommandManagerService cmdMgrService)
        {
            cmdMgrService.AddCommand("Kopíruj Excely", "Zkopíruje podklady z Robota a Toolshopu na plochu", 0, Execute);
        }

        public void Execute()
        {
            try
            {
                SetBarText.Write("Kopíruji Excel soubory na plochu...");

                string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);

                // 1. Lokální soubor
                string localSource = @"M:\FIP_CZ_PRO\2600_Kaizen\99_Zlepsovatelske projekty\2021\2021-030 RPA - Robotic process automation\2021-030 Robotic process automation\2021-030-028 RPA Sklad Třebíč - nastavení stavu\Podklady pro robota.xlsx";
                string localTarget = Path.Combine(desktopPath, "Podklady pro robota.xlsx");

                if (File.Exists(localSource))
                {
                    File.Copy(localSource, localTarget, true);
                    SetBarText.Write("Podklady pro robota úspěšně zkopírovány.");
                }
                else 
                    System.Windows.Forms.MessageBox.Show("Lokální soubor nebyl nalezen:\n" + localSource);
                

                //// 2. SharePoint soubor
                string userRoot = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string spSyncedPath = Path.Combine(userRoot, @"Fraenkische Rohrwerke Gebr. Kirchner GmbH & Co. KG\FIP_CZ_PEEN - Documents\Design Team\Toolshop_drawings.xlsm");
                string spTarget = Path.Combine(desktopPath, "Toolshop_drawings.xlsm");

                if (File.Exists(spSyncedPath))
                {
                    File.Copy(spSyncedPath, spTarget, true);
                    SetBarText.Write("Toolshop úspěšně zkopírován.");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("SharePoint soubor nebyl nalezen:\n" + spSyncedPath);
                }
                MessageBox.Show("Makro dokončeno.");
                SetBarText.Clear();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Chyba při kopírování souborů: " + ex.Message);
                SetBarText.Clear();
            }
        }
    }
}
