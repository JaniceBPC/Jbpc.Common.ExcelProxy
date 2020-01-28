namespace Jbpc.Common.Excel.Proxies
{
    public interface IRange
    {
        void SetText(
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
            int? fontColor = null);
        //void DrawBox(int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin);
        //void FormatGrid(int? gridColor = null, XlBorderWeight borderWeight = XlBorderWeight.xlThin);
        void ApplyAutoFilterToReportBlock();
        IRange Resize(int rowHeight = 1, int colWidth = 1);
        IRange DisplaceAndResize(int displaceRows, int displaceColumns = 0, int rowHeight = 1, int colWidth = 1);
        void CloseWorkbook();
        IRange SetHeadingColumnName(string text, int? columnWidth = null, int orientation = 0);
        int Row { get; }
        int Column { get; }
        int RowCount { get; }
        int ColumnCount { get; }
        IRange Rows { get; }
        IRange Columns { get; }
        object Value2 { get; set; }
    }
}
