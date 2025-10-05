using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Garaoh70.CodeLab.ExcelOpsBench.ExcelOps_CsFW481
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var start = Environment.TickCount;
            MainRoutine(1024);
            var end = Environment.TickCount;
            var elapsed = end - start;
            Console.WriteLine($"ExcelOps_CsFW481: Elapsed {elapsed} ms");
            return;

            #region Local Functions
            void MainRoutine(int count)
            {
                var excel = new Excel.Application();

                // この設定をすると高速化する
                excel.Visible = false;
                excel.ScreenUpdating = false;
                //excel.Calculation = Excel.XlCalculation.xlCalculationManual;
                excel.EnableEvents = false;

                var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                var workbook = excel.Workbooks.Open(Path.Combine(path, @"Sample.xlsx"));
                var worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                var cells = worksheet.Cells;
                var cellA1 = (Excel.Range)cells[1, 1];

                cellA1.Value2 = 0;

                foreach (var _ in Enumerable.Range(1, count))
                {
                    var value = (int)cellA1.Value2 + 1;
                    cellA1.Value2 = value;
                }

                workbook.Close(false);
                excel.Quit();

                Marshal.ReleaseComObject(cellA1);
                Marshal.ReleaseComObject(cells);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
            }
            #endregion
        }
    }
}
