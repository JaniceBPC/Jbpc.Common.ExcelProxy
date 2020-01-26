using Jbpc.Common.Excel.ExtensionMethods;

namespace Jbpc.Common.Excel
{
    public static class WorksheetValues
    {
        public static object[,] GetEntireSheet(string fullyQualifiedWorkbookName, string worksheetName =null)
        {
            using (var reference = ExcelWorkbookWeakReferenceFactory.Instantiate(fullyQualifiedWorkbookName))
            {
                var range = WorksheetUsedRange.GetUsedWorksheetRange(reference.Workbook, worksheetName);

                if (range.Row != 1 || range.Column != 1)
                {
                    var newRows = range.Row - 1 + range.Rows.Count;
                    var newCols = range.Column - 1 + range.Columns.Count;

                    range = range.DisplaceAndResize(-range.Row + 1, -range.Column + 1, newRows, newCols);
                }

                var matrix = (object[,])range.Value2;

                if (ExcelOperations.IsWorkbookAlreadyOpen(fullyQualifiedWorkbookName))
                {
                    range.CloseWorkbook();
                }

                return matrix;
            }
        }
    }
}
