using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.ExtensionMethods
{
    public static class WorksheetExtensionMethods
    {
        public static Range RangeForCell(this Worksheet worksheet, int nthRow, int nthCol)
        {
            return worksheet.Cells[nthRow, nthCol] as Range;
        }
    }
}
