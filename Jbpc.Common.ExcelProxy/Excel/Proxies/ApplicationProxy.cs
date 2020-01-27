using Jbpc.Common.Excel.Proxies.Abstract;
using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.Proxies.Proxies
{
    class ApplicationProxy : IApplication
    {
        private readonly Application application;
        public ApplicationProxy(Application application)
        {
            this.application = application;
        }
        public IWorkbook NewWorkbook => ExcelOperations.NewWorkbook;
    }
}
