using System.Collections.Generic;
using Jbpc.Common.Excel.Proxies.Abstract;

namespace Jbpc.Common.Excel.Proxies
{
    public interface IWorkbook
    {
        List<string> WorksheetNames { get; }
        List<IWorksheet> Worksheets { get; }
        IWorksheet GetWorksheet(string name);
    }
}
