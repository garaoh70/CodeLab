using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeSamples;

internal class Program
{
    static void Main(string[] args)
    {
        var excel = new Excel.Application();
        excel.Visible = false;

        var workbook = excel.Workbooks.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Book1.xlsx"));
        var worksheet = (Excel.Worksheet)workbook.Worksheets[1];

        // 取り消し線のあるセルをクリアする
        var range = worksheet.UsedRange;
        foreach (var item in range.Cells)
        {
            if (item is not Excel.Range cell)
                continue;

            if (cell.Font.Strikethrough is bool strikethrough && strikethrough)
            {
                cell.ClearContents();
            }
            Marshal.ReleaseComObject(item);
        }
        Marshal.ReleaseComObject(range);

        workbook.Close(false);
        excel.Quit();

        Marshal.ReleaseComObject(worksheet);
        Marshal.ReleaseComObject(workbook);
        Marshal.ReleaseComObject(excel);
    }
}
