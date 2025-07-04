using SolidWorks.Interop.sldworks;
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
            var additions = new List<Tuple<string, string, string>>();
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

                // Načtení dat do slovníku
                frame.SetStatusBarText("Opening Excel files...");
                destWB = excelApp.Workbooks.Open(destPath);
                srcWB = excelApp.Workbooks.Open(srcPath, ReadOnly: true);

                var destWS = (Excel.Worksheet)destWB.Sheets[1];
                var srcWS = (Excel.Worksheet)srcWB.Sheets[1];

                int lastRowDest = destWS.Cells[destWS.Rows.Count, DEST_COL_A]
                                         .End(Excel.XlDirection.xlUp).Row;
                int lastRowSrc = srcWS.Cells[srcWS.Rows.Count, SRC_COL_E]
                                         .End(Excel.XlDirection.xlUp).Row;

                frame.SetStatusBarText("Building lookup dictionary...");
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

                // Zápis nových T-čísel a sběr dat pro e-mail
                int additionsCount = 0;
                frame.SetStatusBarText("Processing rows...");
                for (int i = 1; i <= lastRowDest; i++)
                {
                    string partName = Convert.ToString(destWS.Cells[i, DEST_COL_A].Value)?.Trim();
                    string existingT = Convert.ToString(destWS.Cells[i, DEST_COL_F].Value)?.Trim();
                    if (!string.IsNullOrEmpty(partName) && string.IsNullOrEmpty(existingT))
                    {
                        if (srcLookup.TryGetValue(partName, out string newT))
                        {
                            destWS.Cells[i, DEST_COL_F].Value = newT;

                            var destCell = destWS.Cells[i, DEST_COL_F];
                            destCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                            additionsCount++;

                            string author = Convert.ToString(destWS.Cells[i, 3].Value)?.Trim();
                            additions.Add(Tuple.Create(partName, author, newT));
                        }
                    }
                }

                // Uložení změn
                frame.SetStatusBarText("Saving changes...");
                destWB.Save();

                // Odeslání e-mailu přes Outlook
                if (additions.Count > 0)
                {
                    try
                    {
                        var outlookApp = new Outlook.Application();
                        var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                        mailItem.Subject = $"Denní update T-Čísel {DateTime.Now:yyyy-MM-dd}";
                        mailItem.To = recAdresses;         // úprava na reálné adresy
                        var sb = new StringBuilder("Nově přidaná T-Čísla:\r\n\n");
                        // group by author
                        var byAuthor = additions.GroupBy(x => x.Item2);
                        foreach (var group in byAuthor)
                        {
                            // author header
                            sb.AppendLine($"Autor: {group.Key}");

                            // each item under that author, with tabs between fields
                            foreach (var item in group)
                                sb.AppendLine($"{item.Item1}\t\tT-Číslo:{item.Item3}");

                            // separator line between authors
                            sb.AppendLine(new string('-', 55));
                        }

                        //PODPIS
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.AppendLine("S pozdravem.");
                        sb.AppendLine("Váš AUTOKonsturktér.");
                        sb.AppendLine("Fraenkische s.r.o.");

                        mailItem.Body = sb.ToString();
                        mailItem.Display();

                        // Uvolnění COM
                        Marshal.ReleaseComObject(mailItem);
                        Marshal.ReleaseComObject(outlookApp);
                    }
                    catch (Exception mailEx)
                    {
                        File.AppendAllText("Command_LoadTNumbersFromRobot_mail.log", $"{DateTime.Now}: {mailEx}\n");
                    }
                }

                MessageBox.Show($"{additionsCount} new values added to column F.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                File.AppendAllText("Command_LoadTNumbersFromRobot.log", $"{DateTime.Now}: {ex}\n");
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
