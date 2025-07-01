using System.IO;
using System.Windows.Forms;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_6_UpdateLocalExcels : ICommand
    {
        public void Execute()
        {
            string robotSourcePath = @"M:\FIP_CZ_PRO\2600_Kaizen\99_Zlepsovatelske projekty\2021\2021-030 RPA - Robotic process automation\2021-030 Robotic process automation\2021-030-028 RPA Sklad Třebíč - nastavení stavu\Podklady pro robota.xlsx";
            string toolshopSourcePath = @"C:\Users\staffav\Fraenkische Rohrwerke Gebr. Kirchner GmbH & Co. KG\FIP_CZ_PEEN - Documents\Design Team\Toolshop_drawings.xlsm";

            string robotFileName = Path.GetFileName(robotSourcePath);
            string toolshopFileName = Path.GetFileName(toolshopSourcePath);

            string destinationPath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.CodeBase).Replace(@"file:\", string.Empty), @"Resources");

            MessageBox.Show(Path.Combine(destinationPath, robotFileName));
            MessageBox.Show(Path.Combine(destinationPath, toolshopFileName));

            File.Copy(robotSourcePath, Path.Combine(destinationPath, robotFileName), true);
            File.Copy(toolshopSourcePath, Path.Combine(destinationPath, toolshopFileName), true);  
        }   


        public void Register(CommandManagerService cmdMgrService)
        {
            cmdMgrService.AddCommand(
                commandTitle: "Update Local Excels",
                tooltip: "Update local Excel files with latest data",
                iconI: 5, // Use appropriate icon index
                callback: Execute);
        }
    }
}
