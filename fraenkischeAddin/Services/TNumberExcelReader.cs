using System;
using System.Windows.Forms;
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

                bool found = false;

                for (int row = lastRow; row >= 1; row--)
                {

                    string nameCell = sheet.Cells[row, 1].Text as string; // Column A
                    if (!string.IsNullOrWhiteSpace(nameCell) && nameCell.Equals(componentName))
                    {
                        string tNumber = sheet.Cells[row, 6].Text as string; // T-Number from Column A
                        found = true;

                        if (!ConfirmOutputValue(nameCell, tNumber))
                            return null;

                        return string.IsNullOrWhiteSpace(tNumber) ? null : tNumber;
                    }
                }
                if (!found) MessageBox.Show($"Pro dil: {componentName} nebylo nalezeno zadne T-Cislo.");
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

        private bool ConfirmOutputValue(string value, string comp) =>
            MessageBox.Show($"Nalezena shoda v bunce:\n{value}\n{comp}", "Potvrdte nalezenou shodu.", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;


    }
}
