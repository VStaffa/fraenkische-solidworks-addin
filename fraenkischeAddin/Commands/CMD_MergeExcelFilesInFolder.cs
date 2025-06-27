using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_MergeExcelFilesInFolder : ICommand
    {
        public string Title => "Merge Excel Files (BOMs) In Folder.";

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
                commandTitle: Title,
                tooltip: "Excel BOM Merger",
                iconI: 2, // např. 2. ikona ve tvém .bmp
                callback: Execute);
        }
        public void Execute()

        {
            Excel.Application xlApp = null;
            Excel.Workbook outputWorkbook = null;
            Excel.Worksheet outputSheet = null;

            try
            {
                // Create Excel instance
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                // Select folder with FolderBrowserDialog
                string folderPath = SelectFolder();
                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("No folder selected.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (!folderPath.EndsWith("\\"))
                    folderPath += "\\";

                // Create new workbook for output
                outputWorkbook = xlApp.Workbooks.Add();
                outputSheet = (Excel.Worksheet)outputWorkbook.Sheets[1];

                int pasteRow = 1;

                // Process all Excel files in folder
                foreach (string file in Directory.GetFiles(folderPath, "*.xls*"))
                {
                    Excel.Workbook sourceWorkbook = xlApp.Workbooks.Open(file, ReadOnly: true);
                    Excel.Worksheet sourceSheet = (Excel.Worksheet)sourceWorkbook.Sheets[1];

                    // Find last row with data in column A
                    Excel.Range lastCell = sourceSheet.Cells[sourceSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp);
                    int lastRow = lastCell.Row;

                    if (lastRow > 1)
                    {
                        Excel.Range dataRange = sourceSheet.Range[sourceSheet.Cells[1, 1], sourceSheet.Cells[lastRow, sourceSheet.UsedRange.Columns.Count]];
                        Excel.Range destination = outputSheet.Cells[pasteRow, 1];
                        dataRange.Copy(destination);

                        // Update pasteRow for next paste
                        Excel.Range newLastCell = outputSheet.Cells[outputSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp);
                        pasteRow = newLastCell.Row + 1;

                        Marshal.ReleaseComObject(dataRange);
                        Marshal.ReleaseComObject(destination);
                        Marshal.ReleaseComObject(newLastCell);
                    }

                    sourceWorkbook.Close(false);
                    Marshal.ReleaseComObject(sourceSheet);
                    Marshal.ReleaseComObject(sourceWorkbook);
                }

                // Save merged workbook
                string savePath = Path.Combine(folderPath, "Spojeny_BOM_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                outputWorkbook.SaveAs(savePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                MessageBox.Show($"Merge completed and saved to:\n{savePath}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

                outputWorkbook.Close(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (outputSheet != null) Marshal.ReleaseComObject(outputSheet);
                if (outputWorkbook != null) Marshal.ReleaseComObject(outputWorkbook);
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private string SelectFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder with Excel files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() == DialogResult.OK)
                    return dialog.SelectedPath;
                else
                    return null;
            }
        }

    }
}
