using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_4_LoadTNumbersFromRobot : ICommand
    {
        private readonly SldWorks _swApp;

        // Seznam email adres 
        private readonly string recAdresses = "vaclav.staffa@fraenkische-cz.com;" +
            "tomas.kalina@fraenkische-cz.com;" +
            "jaroslav.hruska@fraenkische-cz.com;" +
            "jaromir.hroch@fraenkische-cz.com;" +
            "lubos.hromadko@fraenkische-cz.com;" +
            "jiri.kalis@fraenkische-cz.com;" +
            "zdenek.sveda@fraenkische-cz.com";

        // Konstanty pro sloupce a filtr
        private const int DEST_COL_A = 1;
        private const int DEST_COL_F = 6;
        private const int SRC_COL_A = 1;
        private const int SRC_COL_E = 5;
        private const string EXCEL_FILE_FILTER = "Excel Files|*.xlsx;*.xlsm;*.xls";

        public CMD_4_LoadTNumbersFromRobot(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "Load T-Numbers From Robot";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(Title, "Load New Daily T-Numbers From Robot", 3, Execute);
        }

        public void Execute()
        {
            // Rozšířený seznam o stavy pro part i drawing
            var additions = new List<(string PartName, string Author, string TNumber,
                                      bool FileFound, bool PropertyAdded,
                                      bool DrawingFound, bool DrawingSaved)>();
            IFrame frame = _swApp.Frame();

            // Výběr souborů
            var ofd = new OpenFileDialog { Title = "Select 'TOOL_SHOP' Excel file", Filter = EXCEL_FILE_FILTER };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            string destPath = ofd.FileName;

            ofd.Title = "Select 'Podklady pro Robota' Excel file";
            if (ofd.ShowDialog() != DialogResult.OK) return;
            string srcPath = ofd.FileName;

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook destWB = null, srcWB = null;

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                frame.SetStatusBarText("Opening Excel files.");
                destWB = excelApp.Workbooks.Open(destPath);
                srcWB = excelApp.Workbooks.Open(srcPath, ReadOnly: true);

                var destWS = (Excel.Worksheet)destWB.Sheets[1];
                var srcWS = (Excel.Worksheet)srcWB.Sheets[1];

                int lastRowDest = destWS.Cells[destWS.Rows.Count, DEST_COL_A]
                                         .End(Excel.XlDirection.xlUp).Row;
                int lastRowSrc = srcWS.Cells[srcWS.Rows.Count, SRC_COL_E]
                                         .End(Excel.XlDirection.xlUp).Row;

                frame.SetStatusBarText("Building lookup dictionary.");
                var srcLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 1; j <= lastRowSrc; j++)
                {
                    string srcE = Convert.ToString(srcWS.Cells[j, SRC_COL_E].Value)?.Trim();
                    string srcA = Convert.ToString(srcWS.Cells[j, SRC_COL_A].Value)?.Trim();
                    if (!string.IsNullOrEmpty(srcE) && !string.IsNullOrEmpty(srcA))
                        foreach (var token in srcE.Split(' '))
                            if (!srcLookup.ContainsKey(token))
                                srcLookup[token] = srcA;
                }

                int additionsCount = 0;
                frame.SetStatusBarText("Processing rows.");
                for (int i = 1; i <= lastRowDest; i++)
                {
                    string partName = Convert.ToString(destWS.Cells[i, DEST_COL_A].Value)?.Trim();
                    string existingT = Convert.ToString(destWS.Cells[i, DEST_COL_F].Value)?.Trim();
                    if (!string.IsNullOrEmpty(partName) && string.IsNullOrEmpty(existingT))
                    {
                        if (srcLookup.TryGetValue(partName, out string newT))
                        {
                            // Zápis do Excelu
                            destWS.Cells[i, DEST_COL_F].Value = newT;
                            destWS.Cells[i, DEST_COL_F].Interior.Color =
                                System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                            additionsCount++;

                            string author = Convert.ToString(destWS.Cells[i, 3].Value)?.Trim();

                            // Otevření a úprava part modelu
                            string partDir = @"C:\Users\staff\Desktop\BFP";
                            string partPath = Path.Combine(partDir, partName, partName + ".sldprt");

                            bool fileFound = File.Exists(partPath);
                            bool propertyAdded = false;
                            bool drawingFound = false;
                            bool drawingSaved = false;

                            if (fileFound)
                            {
                                var model = _swApp.OpenDoc6(
                                    partPath,
                                    (int)swDocumentTypes_e.swDocPART,
                                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                                    "", 0, 0) as ModelDoc2;
                                if (model != null)
                                {
                                    var cusMgr = model.Extension.CustomPropertyManager[""];
                                    cusMgr.Add3("T-Number",
                                               (int)swCustomInfoType_e.swCustomInfoText,
                                               newT, 1);
                                    model.ForceRebuild3(true);
                                    model.Save();
                                    propertyAdded = true;
                                    _swApp.CloseDoc(model.GetTitle());

                                    // Otevření a úprava výkresu
                                    string drwPath = Path.Combine(partDir, partName, partName + ".slddrw");
                                    drawingFound = File.Exists(drwPath);
                                    if (drawingFound)
                                    {
                                        var drw = _swApp.OpenDoc6(
                                            drwPath,
                                            (int)swDocumentTypes_e.swDocDRAWING,
                                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                                            "", 0, 0) as ModelDoc2;
                                        if (drw != null)
                                        {
                                            drw.ForceRebuild3(true);
                                            drw.Save();
                                            drawingSaved = true;
                                            _swApp.CloseDoc(drw.GetTitle());
                                        }
                                    }
                                }
                            }

                            additions.Add((partName, author, newT,
                                           fileFound, propertyAdded,
                                           drawingFound, drawingSaved));
                        }
                    }
                }

                // Uložení Excelu
                frame.SetStatusBarText("Saving changes.");
                destWB.Save();

                // Odeslání e-mailu
                if (additions.Count > 0)
                {
                    try
                    {
                        var outlookApp = new Outlook.Application();
                        var mailItem = (Outlook.MailItem)outlookApp.CreateItem(
                            Outlook.OlItemType.olMailItem);

                        mailItem.Subject = $"Denní update T-Čísel {DateTime.Now:yyyy-MM-dd}";
                        mailItem.To = recAdresses;
                        var sb = new StringBuilder("Nově přidaná T-Čísla:\r\n\n");

                        var byAuthor = additions.GroupBy(x => x.Author);
                        foreach (var group in byAuthor)
                        {
                            sb.AppendLine($"Autor: {group.Key}");
                            foreach (var item in group)
                            {
                                sb.AppendLine(
                                    $"DÍL: {item.PartName}\t" +
                                    $"T-Číslo: {item.TNumber}\t\t" +
                                    $"Díl nalezen: {(item.FileFound ? "OK" : "Nenalezen")}\t\t" +
                                    $"Díl upraven: {(item.PropertyAdded ? "OK" : "Chyba")}\t\t" +
                                    $"Výkres nalezen: {(item.DrawingFound ? "OK" : "Nenalezen")}\t\t" +
                                    $"Výkres upraven: {(item.DrawingSaved ? "OK" : "Chyba")}"
                                );
                            }
                            sb.AppendLine(new string('-', 80));
                        }

                        sb.AppendLine();
                        sb.AppendLine("S pozdravem,");
                        sb.AppendLine("Váš AUTOKonstrukter");
                        sb.AppendLine("Fraenkische s.r.o.");

                        mailItem.Body = sb.ToString();
                        mailItem.Display();

                        Marshal.ReleaseComObject(mailItem);
                        Marshal.ReleaseComObject(outlookApp);
                    }
                    catch (Exception mailEx)
                    {
                        File.AppendAllText(
                            "Command_LoadTNumbersFromRobot_mail.log",
                            $"{DateTime.Now}: {mailEx}\n");
                    }
                }

                MessageBox.Show(
                    $"{additionsCount} new values added to column F.",
                    "Done",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                File.AppendAllText(
                    "Command_LoadTNumbersFromRobot.log",
                    $"{DateTime.Now}: {ex}\n");
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                srcWB?.Close(false);
                destWB?.Close(true);
                excelApp.DisplayAlerts = true;
                excelApp.ScreenUpdating = true;
                excelApp.StatusBar = false;
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                frame.SetStatusBarText("Ready");
            }
        }
    }
}
