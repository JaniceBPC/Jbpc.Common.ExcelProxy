using Jbpc.Common.Excel.ExtensionMethods;
using Microsoft.Office.Interop.Excel;
//using System.Reactive.Disposables;

namespace Jbpc.Common.Excel.Proxies
{
    public class RangeProxy : IRange
    {
        //private CompositeDisposable gg;

        private readonly Range range;
        public RangeProxy(Range range)
        {
            this.range = range;
        }
        public void SetText(
            string text,
            bool? wrapText = null,
            //XlHAlign xlHAligh = XlHAlign.xlHAlignLeft,
            //XlVAlign xlVAlign = XlVAlign.xlVAlignCenter,
            float orientation = 0,
            float indetLevel = 0,
            float fontSize = 12,
            bool bold = false,
            bool? mergeCells = null,
            int? interiorColor = null,
            int? fontColor = null) => range.SetText(text, wrapText, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignCenter, orientation, indetLevel, fontSize, bold, mergeCells, interiorColor, fontColor);

        public void SetText(
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
            bool? mergeCells = null,
            int? interiorColor = null,
            int? fontColor = null) => range.SetText(row, col, text, wrapText, xlHAligh, xlVAlign, orientation, indetLevel, fontSize, bold, mergeCells, interiorColor, fontColor);

        public void DrawBox(int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin) =>
            range.DrawBox(gridColor, borderWeight);

        public void FormatGrid(int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin) =>
            range.FormatGrid(gridColor, borderWeight);

        public void ApplyAutoFilterToReportBlock() => range.ApplyAutoFilterToReportBlock();

        public IRange Resize(int rowHeight = 1, int colWidth = 1) => new RangeProxy(range.Resize(rowHeight, colWidth));

        public IRange DisplaceAndResize(int displaceRows, int displaceColumns = 0, int rowHeight = 1, int colWidth = 1) =>
            new RangeProxy(range.DisplaceAndResize(displaceRows, displaceColumns, rowHeight, colWidth));

        public void CloseWorkbook() => range.CloseWorkbook();

        public IRange SetHeadingColumnName(string text, int? columnWidth = null, int orientation = 0) => new RangeProxy(range.SetHeadingColumnName(text, columnWidth, orientation));
        public int Row => range.Row;
        public int Column => range.Column;
        public int I => range.Column;
        public IRange Rows => new RangeProxy(range.Rows);
        public IRange Columns => new RangeProxy(range.Columns);
        public int RowCount => range.Rows.Count;
        public int ColumnCount => range.Columns.Count;
        public object Value2
        {
            get => range.Value2;
            set => range.Value2 = value;
        }

    }
}
