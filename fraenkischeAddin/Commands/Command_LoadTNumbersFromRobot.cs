using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.sldworks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fraenkische.SWAddin.Commands
{
    public class Command_LoadTNumbersFromRobot : ICommand
    {
        private readonly SldWorks _swApp;

        // Recommendation 1: Use constants for column indices and file filter
        private const int DEST_COL_A = 1;
        private const int DEST_COL_F = 6;
        private const int SRC_COL_A = 1;
        private const int SRC_COL_E = 5;
        private const string EXCEL_FILE_FILTER = "Excel Files|*.xlsx;*.xlsm;*.xls";

        public Command_LoadTNumbersFromRobot(SldWorks swApp)
        {
            _swApp = swApp;
        }
        public string Title => "Load T-Numbers From Robot";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(Title, "Load New Daily T-Numbers From Robot", 1, Execute);
        }

        public void Execute()
        {

            IFrame frame;
            frame = _swApp.Frame();

            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Select 'TOOLBOX' Excel file",
                Filter = EXCEL_FILE_FILTER
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;
            string destPath = ofd.FileName;


            frame.SetStatusBarText("Loading 'TOOLBOX' Excel file...");
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

                frame.SetStatusBarText("Opening Excel files...");
                destWB = excelApp.Workbooks.Open(destPath);
                srcWB = excelApp.Workbooks.Open(srcPath, ReadOnly: true);

                Excel.Worksheet destWS = destWB.Sheets[1];
                Excel.Worksheet srcWS = srcWB.Sheets[1];

                int lastRowDest = destWS.Cells[destWS.Rows.Count, DEST_COL_A].End(Excel.XlDirection.xlUp).Row;
                int lastRowSrc = srcWS.Cells[srcWS.Rows.Count, SRC_COL_E].End(Excel.XlDirection.xlUp).Row;

                int additionsCount = 0;

                frame.SetStatusBarText("Building lookup dictionary...");
                var srcLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 1; j <= lastRowSrc; j++)
                {
                    string srcE = Convert.ToString(srcWS.Cells[j, SRC_COL_E].Value)?.Trim();
                    string srcA = Convert.ToString(srcWS.Cells[j, SRC_COL_A].Value);
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

                    frame.SetStatusBarText("Processing rows...");
                    for (int i = 1; i <= lastRowDest; i++)
                    {
                        string destA = Convert.ToString(destWS.Cells[i, DEST_COL_A].Value)?.Trim();
                        string destF = Convert.ToString(destWS.Cells[i, DEST_COL_F].Value)?.Trim();

                        if (!string.IsNullOrEmpty(destA) && string.IsNullOrEmpty(destF))
                        {
                            if (srcLookup.TryGetValue(destA, out string srcA))
                            {
                                destWS.Cells[i, DEST_COL_F].Value = srcA;
                                destWS.Cells[i, DEST_COL_F].Interior.Color = ColorTranslator.ToOle(Color.Orange);
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

                frame.SetStatusBarText("Saving changes...");
                destWB.Save();
                MessageBox.Show($"{additionsCount} new values added to column F.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                frame.SetStatusBarText("Ready");
            }
        }
    }


    // Simple progress bar form for recommendation 5
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private int max;

        public ProgressForm(int max)
        {
            this.max = max;
            this.Text = "Processing...";
            this.Width = 400;
            this.Height = 80;
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = max
            };
            Controls.Add(progressBar);
        }

        public void UpdateProgress(int value)
        {
            if (value > max) value = max;
            progressBar.Value = value;
        }
    }
}