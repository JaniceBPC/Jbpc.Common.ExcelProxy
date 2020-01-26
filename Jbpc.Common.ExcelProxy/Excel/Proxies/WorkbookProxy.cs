using System.Collections.Generic;
using System.Linq;
using Jbpc.Common.Excel.ExtensionMethods;
using Jbpc.Common.Excel.Proxies.Abstract;
using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.Proxies
{
    public class WorkbookProxy : IWorkbook
    {
        private readonly Workbook workbook;
        public WorkbookProxy(Workbook workbook)
        {
            this.workbook = workbook;
        }
        public List<string> WorksheetNames => workbook.WorksheetNames();
        public List<IWorksheet> Worksheets => workbook.WorksheetsList().Select(x=> new WorksheetProxy(x)).Cast<IWorksheet>().ToList();
    }
}
 