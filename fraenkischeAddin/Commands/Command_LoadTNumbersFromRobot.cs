using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using Excel = Microsoft.Office.Interop.Excel;



namespace Fraenkische.SWAddin.Commands
{
    public class Command_LoadTNumbersFromRobot : ICommand
    {
        private readonly ISldWorks _swApp;

        public Command_LoadTNumbersFromRobot(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public string Title => "LoadTNumbersFromRobot";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(Title, "Load New Daily TNumbers From Robot", 1, Execute);
        }

        public void Execute()
        {

            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Select destination Excel file",
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls"
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;
            string destPath = ofd.FileName;

            ofd.Title = "Select source Excel file";
            if (ofd.ShowDialog() != DialogResult.OK) return;
            string srcPath = ofd.FileName;

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook destWB = null;
            Excel.Workbook srcWB = null;

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                destWB = excelApp.Workbooks.Open(destPath);
                srcWB = excelApp.Workbooks.Open(srcPath, ReadOnly: true);

                Excel.Worksheet destWS = destWB.Sheets[1];
                Excel.Worksheet srcWS = srcWB.Sheets[1];

                int lastRowDest = destWS.Cells[destWS.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row;
                int lastRowSrc = srcWS.Cells[srcWS.Rows.Count, "E"].End(Excel.XlDirection.xlUp).Row;

                int additionsCount = 0;

                for (int i = 1; i <= lastRowDest; i++)
                {
                    string destA = Convert.ToString(destWS.Cells[i, 1].Value)?.Trim();
                    string destF = Convert.ToString(destWS.Cells[i, 6].Value)?.Trim();

                    MessageBox.Show(destA + "  " + destF);

                    if (!string.IsNullOrEmpty(destA) && string.IsNullOrEmpty(destF))
                    {
                        for (int j = lastRowSrc; j >= 1; j--)
                        {
                            string srcE = Convert.ToString(srcWS.Cells[j, 5].Value)?.Trim();
                            if (!string.IsNullOrEmpty(srcE))
                            {
                                string[] tokens = srcE.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (var token in tokens)
                                {
                                    if (token == destA)
                                    {
                                        string srcA = Convert.ToString(srcWS.Cells[j, 1].Value);
                                        destWS.Cells[i, 6].Value = srcA;
                                        destWS.Cells[i, 6].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                                        additionsCount++;
                                        goto NextRow;
                                    }
                                }
                            }
                        }
                    }
                NextRow:;
                    if (i % 50 == 0)
                        excelApp.StatusBar = $"Processing row {i} of {lastRowDest}";
                }

                destWB.Save();
                MessageBox.Show($"{additionsCount} new values added to column F.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
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
            }
        }

        }
}
