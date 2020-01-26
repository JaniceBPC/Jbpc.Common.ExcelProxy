using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.ExtensionMethods
{
    public static class RangeExtensionMethods
    {
        public static int GridColor => ExcelConstants.GridColor;

        public static void SetText(
            this Range rngHome,
            string text,
            bool? wrapText = null,
            XlHAlign xlHAligh = XlHAlign.xlHAlignLeft,
            XlVAlign xlVAlign = XlVAlign.xlVAlignCenter,
            float orientation = 0,
            float indetLevel = 0,
            float fontSize = 12,
            bool bold = false,
            bool? mergeCells = null,
            int? interiorColor = null,
            int? fontColor = null)
        {
            rngHome.Value2 = text;
            rngHome.HorizontalAlignment = xlHAligh;
            rngHome.VerticalAlignment = xlVAlign;
            rngHome.Font.Size = fontSize;
            rngHome.Font.Bold = bold;
            rngHome.IndentLevel = indetLevel;
            rngHome.Orientation = orientation;
            rngHome.WrapText = wrapText;

            if (wrapText.HasValue) rngHome.WrapText = wrapText;
            if (interiorColor.HasValue) rngHome.Interior.Color = interiorColor.Value;
            if (fontColor.HasValue) rngHome.Font.Color = fontColor.Value;
            if (mergeCells.HasValue && mergeCells.Value) rngHome.Merge();
        }
        public static void SetText(
            this Range rngHome,
            int row,
            int col,
            string text,
            bool wrapText = false,
            XlHAlign xlHAligh = XlHAlign.xlHAlignLeft,
            XlVAlign xlVAlign = XlVAlign.xlVAlignCenter,
            float orientation = 0,
            float indetLevel = 0,
            float fontSize = 12,
            bool bold = false,
            bool? mergeCells = false,
            int? interiorColor = null,
            int? fontColor = null)
        {
            var rng = (Range) rngHome.Cells[row, col];

            rng.SetText(text, wrapText, xlHAligh, xlVAlign, orientation, indetLevel, fontSize, bold, mergeCells, interiorColor, fontColor);

        }
        public static void DrawBox(this Range rngHome, int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin)
        {
            var c = gridColor ?? GridColor;
            var w = borderWeight;

            var x = rngHome.Borders;

            M(x[XlBordersIndex.xlEdgeLeft], c, w);
            M(x[XlBordersIndex.xlEdgeBottom], c, w);
            M(x[XlBordersIndex.xlEdgeRight], c, w);
            M(x[XlBordersIndex.xlEdgeTop], c, w);
        }
        public static void FormatGrid(this Range rngHome, int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin)
        {
            var c = gridColor ?? GridColor;
            var w = borderWeight;

            var x = rngHome.Borders;

            M(x[XlBordersIndex.xlEdgeLeft], c, w);
            M(x[XlBordersIndex.xlEdgeBottom], c, w);
            M(x[XlBordersIndex.xlEdgeRight], c, w);
            M(x[XlBordersIndex.xlEdgeTop], c, w);
            M(x[XlBordersIndex.xlInsideHorizontal], c, w);
            M(x[XlBordersIndex.xlInsideVertical], c, w);
        }
        private static void M(Border borders, int gridColor, XlBorderWeight borderWeight)
        {
            borders.LineStyle = XlLineStyle.xlContinuous;
            borders.Weight = borderWeight;
            borders.Color = gridColor;
        }
        public static void ApplyAutoFilterToReportBlock(this Range rngBlock)
        {
            System.Diagnostics.Debug.Assert(rngBlock.Row>1);

            var rngHeader = rngBlock.get_Offset(-1).get_Resize(rngBlock.Rows.Count + 1, rngBlock.Columns.Count);

            rngHeader.AutoFilter(1, Operator: XlAutoFilterOperator.xlFilterValues, VisibleDropDown: true);
        }
        public static Range Resize(this Range range, int rowHeight = 1, int colWidth = 1)
        {
            var msg1 = range.A1_A1();

            var newRange = range.get_Resize(rowHeight, colWidth);

            var msg2 = range.A1_A1();

            return newRange;
        }
        public static Range DisplaceAndResize(this Range range, int displaceRows, int displaceColumns = 0,
            int rowHeight = 1, int colWidth = 1)
        {
            var msg1 = range.A1_A1();

            if (range.Row + displaceRows < 1) System.Diagnostics.Debugger.Break();

            var newRange = range.get_Offset(displaceRows, displaceColumns).get_Resize(rowHeight, colWidth);

            var msg2 = range.A1_A1();

            return newRange;
        }
        public static void CloseWorkbook(this Range range)
        {
            var workbook = (Workbook) range.Worksheet.Parent;

            workbook.Close();
        }
        public static Range SetHeadingColumnName(this Range rngRange, string text, int? columnWidth = null, int orientation = 0)
        {
            rngRange.SetText(
                text,
                xlVAlign: XlVAlign.xlVAlignBottom,
                xlHAligh: XlHAlign.xlHAlignCenter,
                wrapText: true,
                fontSize: 10,
                bold: true,
                orientation: orientation,
                interiorColor: ExcelConstants.HeaadingsBlue,
                fontColor: ExcelConstants.White);

            if (columnWidth.HasValue)
            {
                rngRange.EntireColumn.ColumnWidth = columnWidth.Value;
            }

            return rngRange.DisplaceAndResize(0, 1);
        }
    }
}
