using System;
using System.Collections.Generic;
using System.Linq;
using Jbpc.Common.Excel.ExtensionMethods;
using Jbpc.Common.Excel.Proxies.Abstract;
using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.Proxies
{
    public class WorkbookProxy : IWorkbook
    {
        private Workbook workbook;
        public WorkbookProxy(Workbook workbook)
        {
            this.workbook = workbook;
        }
        public IWorksheet GetWorksheet(string name) => new WorksheetProxy(workbook.GetWorksheet(name));
        public List<string> WorksheetNames => workbook.WorksheetNames();
        public List<IWorksheet> Worksheets =>
                workbook.WorksheetsList()
                    .Select(x => new WorksheetProxy(x))
                    .Cast<IWorksheet>()
                    .ToList();
        public void Close()
        {
            workbook.Close();

            workbook = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
 