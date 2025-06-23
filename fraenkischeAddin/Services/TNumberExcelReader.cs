using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fraenkische.SWAddin.Services
{
    public class TNumberExcelReader
    {
        private readonly string _excelPath;

        public TNumberExcelReader(string excelPath)
        {
            _excelPath = excelPath;
        }

        public string GetTNumberForComponent(string componentName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                workbook = xlApp.Workbooks.Open(_excelPath, ReadOnly: true);
                Excel.Worksheet sheet = workbook.Sheets[1];
                Excel.Range usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;

                for (int row = lastRow; row >= 1; row--)
                {
                    string nameCell = sheet.Cells[row, 5].Text as string;
                    if (!string.IsNullOrWhiteSpace(nameCell) &&
                        nameCell.IndexOf(componentName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string tNumber = sheet.Cells[row, 1].Text as string;
                        return string.IsNullOrWhiteSpace(tNumber) ? null : tNumber;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Excel read error: " + ex.Message);
            }
            finally
            {
                workbook?.Close(false);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }

            return null;
        }
    }
}
