using System;
using System.IO;

namespace Jbpc.Common.Excel
{
    public static class ExcelWorkbookWeakReferenceFactory
    {
        public static ExcelWorkbookWeakReference Instantiate(string path)
        {
            if (!File.Exists(path)) throw new ApplicationException($"Workbook filename does not exist={path}");

            var reference = new ExcelWorkbookWeakReference()
            {
                Path = path,
                IsAlreadyOpened = ExcelOperations.IsWorkbookAlreadyOpen(path),
                WorkbookName = ExcelOperations.WorkbookName(path)
            };

            if (reference.IsAlreadyOpened)
            {
                reference.Workbook = ExcelOperations.GetAlreadyOpenedWorkbook(path);

                return reference;
            }

            reference.Workbook = ExcelOperations.OpenWorkbook(path);

            return reference;
        }
    }
}
