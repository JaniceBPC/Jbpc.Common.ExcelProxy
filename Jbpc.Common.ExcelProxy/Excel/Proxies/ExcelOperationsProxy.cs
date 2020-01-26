namespace Jbpc.Common.Excel.Proxies
{
    public static class ExcelOperationsProxy
    {
        public static IWorkbook NewWorkbook => new WorkbookProxy(ExcelOperations.NewWorkbook);
    }
}
