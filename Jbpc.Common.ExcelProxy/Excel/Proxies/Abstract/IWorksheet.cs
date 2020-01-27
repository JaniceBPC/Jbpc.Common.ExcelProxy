namespace Jbpc.Common.Excel.Proxies.Abstract
{
    public interface IWorksheet
    {
        string Name { get; set; }
        IRange RangeForCell(int nthRow, int nthCol);
        IRange UsedRange();
    }
}
