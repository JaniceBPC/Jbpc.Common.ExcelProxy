using System;
using Jbpc.Common.Excel.Proxies;

namespace Jbpc.Common.Excel
{
    public class ExcelWorkbookWeakReference : IDisposable
    {
        public bool IsAlreadyOpened { get; set; }
        public IWorkbook Workbook { get; set;  }
        public string WorkbookName { get; set;  }
        public string Path { get; set; }
        public void Dispose()
        {
            Dispose(true);
        }

        public void Dispose(bool flag)
        {
            if (!IsAlreadyOpened)
            {
                Workbook?.Close();
            }
        }
    }
}
