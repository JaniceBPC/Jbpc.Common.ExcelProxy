using Jbpc.Common.Excel.Proxies.Abstract;
using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.Proxies
{
    public class WorksheetProxy : IWorksheet
    {
        private readonly Worksheet worksheet;

        public WorksheetProxy(Worksheet worksheet)
        {
            this.worksheet = worksheet;
        }
        public string Name
        {
            get => worksheet.Name;
            set => worksheet.Name = value;
        }
    }
}
