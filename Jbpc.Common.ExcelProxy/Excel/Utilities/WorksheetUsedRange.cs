using System;
using System.Linq;
using Jbpc.Common.Excel.Proxies;
using IRange = Jbpc.Common.Excel.Proxies.IRange;

namespace Jbpc.Common.Excel
{
    public static class WorksheetUsedRange
    {
        public static IRange GetUsedWorksheetRange(IWorkbook workbook, string worksheetName = null)
        {
            worksheetName = worksheetName ?? workbook.WorksheetNames.Last();

            var worksheet = workbook.GetWorksheet(worksheetName);

            if (worksheet == null) throw new ApplicationException($"Failed to open worksheet={worksheetName}, in workbook={worksheetName}");

            return worksheet.UsedRange();
        }

    }
}
