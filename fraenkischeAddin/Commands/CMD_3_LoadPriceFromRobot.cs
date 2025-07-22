using Fraenkische.SWAddin.Core;
using Fraenkische.SWAddin.Services;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_3_LoadPriceFromRobot : ICommand
    {
        // Recommendation 1: Use constants for column indices and file filter
        private const int DEST_COL_A = 1;
        private const int DEST_COL_D = 5;
        private const int SRC_COL_O = 15;
        private const int SRC_COL_E = 5;
        private const string EXCEL_FILE_FILTER = "Excel Files|*.xlsx;*.xlsm;*.xls";

        public string Title => "Load PRICE from ROBOT";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(Title, "Load Prices of found parts.", 2, Execute);
        }

        public void Execute()
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Select 'DESTINANTION' Excel file",
                Filter = EXCEL_FILE_FILTER
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;
            string destPath = ofd.FileName;

            ofd.Title = "Select 'Podklady pro Robota' Excel file";
            if (ofd.ShowDialog() != DialogResult.OK) return;
            string srcPath = ofd.FileName;

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook destWB = null;
            Excel.Workbook srcWB = null;

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                SetBarText.Write("Opening Excel files..."); 
                destWB = excelApp.Workbooks.Open(destPath);
                srcWB = excelApp.Workbooks.Open(srcPath, ReadOnly: true);

                Excel.Worksheet destWS = destWB.Sheets[1];
                Excel.Worksheet srcWS = srcWB.Sheets[1];

                int lastRowDest = destWS.Cells[destWS.Rows.Count, DEST_COL_A].End(Excel.XlDirection.xlUp).Row;
                int lastRowSrc = srcWS.Cells[srcWS.Rows.Count, SRC_COL_E].End(Excel.XlDirection.xlUp).Row;

                int additionsCount = 0;

                SetBarText.Write("Building lookup dictionary...");
                var srcLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 1; j <= lastRowSrc; j++)
                {
                    string srcE = Convert.ToString(srcWS.Cells[j, SRC_COL_E].Value)?.Trim();
                    string srcA = Convert.ToString(srcWS.Cells[j, SRC_COL_O].Value);
                    if (!string.IsNullOrEmpty(srcE) && !string.IsNullOrEmpty(srcA))
                    {
                        string[] tokens = srcE.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var token in tokens)
                        {
                            if (!srcLookup.ContainsKey(token))
                                srcLookup[token] = srcA;
                        }
                    }
                }

                using (var progress = new ProgressForm(lastRowDest))
                {
                    progress.Show();
                    progress.UpdateProgress(0);

                    SetBarText.Write("Processing rows...");
                    for (int i = 1; i <= lastRowDest; i++)
                    {
                        string destA = Convert.ToString(destWS.Cells[i, DEST_COL_A].Value)?.Trim();
                        string destF = Convert.ToString(destWS.Cells[i, DEST_COL_D].Value)?.Trim();

                        if (!string.IsNullOrEmpty(destA) && string.IsNullOrEmpty(destF))
                        {
                            if (srcLookup.TryGetValue(destA, out string srcA))
                            {
                                destWS.Cells[i, DEST_COL_D].Value = srcA;
                                destWS.Cells[i, DEST_COL_D].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                                additionsCount++;
                            }
                        }

                        if (i % 10 == 0 || i == lastRowDest)
                        {
                            progress.UpdateProgress(i);
                            Application.DoEvents();
                        }
                    }
                }

                SetBarText.Write("Saving changes...");
                destWB.Save();
                MessageBox.Show($"{additionsCount} new values added to column D.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText("Command_LoadTNumbersFromRobot.log", $"{DateTime.Now}: {ex}\n");
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
                SetBarText.Write("Ready");
            }
        }
    }
}