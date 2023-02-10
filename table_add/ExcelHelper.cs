using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace otchet_fill
{
    class ExcelHelper: IDisposable 
    {
        private Excel.Application _excel;
        private Workbook _workbook;

        public ExcelHelper() => _excel = new Excel.Application();
        internal bool Open(string filePath)
        {
            try
            {
                _workbook = _excel.Workbooks.Open(filePath);
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return false;
        }
        internal bool AutoFit()
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Columns.EntireColumn.AutoFit();
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
            return true;
        }
        internal bool Add(string filePath)
        {
            try
            {
                _workbook = _excel.Workbooks.Add();
                _workbook.SaveAs(filePath);
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return false;
        }
        
        internal int FindLastRow()
        {
            try
            {
                int i = 1;
                while (((Excel.Worksheet)_excel.ActiveSheet).Cells[i, 1].Value != null)
                {
                    i++;
                }
                return i;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return 2;
        }
        internal bool Set(int row, int col, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, col] = data;
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return false;
        }

        internal void Save()
        {
            try
            {
                    _workbook.Save();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void Dispose()
        {
            try
            {
                _excel.Quit();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        
    }
}
