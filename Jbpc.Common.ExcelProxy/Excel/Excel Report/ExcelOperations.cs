using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using Jbpc.Common.Excel.Proxies;

namespace Jbpc.Common.Excel
{
    public static class ExcelOperations
    {
        public static void PerformWithLock(System.Action action)
        {
            lock (ExcelApplication)
            {
                action();
            }
        }
        private static Application excelApplication;

        public static Application ExcelApplication
        {
            get
            {
                if (excelApplication == null)
                {
                    excelApplication = OfficeApplication.Singleton.Application;
                }

                return excelApplication;
            }
        }

        public static IWorkbook OpenWorkbook(string workbookFilename, bool readOnly = true)
        {
            if (!File.Exists(workbookFilename))
            {
                throw new ApplicationException($"Open workbook - workbook not specified: {workbookFilename}");
            }
            return OpenWorkbookWithRetry(workbookFilename, readOnly: readOnly);

        }
        public static IWorkbook NewWorkbook => new WorkbookProxy(ExcelApplication.Workbooks.Add());
        public static IWorkbook OpenReportWorkbook(string templateName)
        {

            if (templateName == null)
            {
                throw new ApplicationException($"Open workbook - workbook name is null: {templateName}");
            }

            var workbookFilename = Path.Combine(TemplateWorkbookDirectory.PathName(), templateName);

            if (!File.Exists(workbookFilename))
            {
                throw new ApplicationException($"Open workbook - workbook not specified: {workbookFilename}");
            }

            return OpenWorkbookWithRetry(workbookFilename);
        }
        public static IWorkbook OpenWorkbookWithRetry(string path, bool closeThenReopen = false, bool readOnly = true)
        {
            var workbook = GetAlreadyOpenedWorkbook(path);

            if (workbook != null && !closeThenReopen)
            {
                return workbook; 
            }

            PerformOperationWithRecovery.PerformOperation(() =>
            {
                workbook?.Close();

                try
                {
                    var xlWorkbook = ExcelApplication.Workbooks.Open(path, readOnly);

                    workbook = new WorkbookProxy(xlWorkbook);
                }
                catch (Exception)
                {
                    throw new ApplicationException($"Failed to open workbook=\"{path}\"");
                }
            });
            return workbook;
        }
        public static IWorkbook GetWorkbook(string workbookName, string path)
        {
            IWorkbook workbook = null;

            PerformOperationWithRecovery.PerformOperation(() =>
            {
                ExcelApplication.Workbooks.Open(path, ReadOnly: true);

                try
                {
                    workbook = new WorkbookProxy(ExcelApplication.Workbooks.Open(path, ReadOnly: true)); 
                }
                catch (Exception)
                {
                    throw new ApplicationException($"Failed to open workbook={workbookName} path=\"{path}\"");
                }
            });

            return workbook;
        }
        public static string WorkbookName(string workbookFileName)
        {
            return workbookFileName.Contains(".") ?
                Path.GetFileNameWithoutExtension(workbookFileName) :
                workbookFileName;
        }
        public static bool IsWorkbookAlreadyOpen(string workbookName)
        {
            return GetAlreadyOpenedWorkbook(WorkbookName(workbookName)) != null;
        }
        public static IWorkbook GetAlreadyOpenedWorkbook(string path)
        {
            string workbookName = WorkbookName(path);

            foreach (Workbook each in ExcelApplication.Workbooks)
            {
                if (each.Name == workbookName)
                {
                    return new WorkbookProxy(each);
                }
            }

            return null;
        }
        public static void MakeWorkbookVisible(bool visible) => PerformOperationWithRecovery.PerformOperation(() => ExcelApplication.Visible = visible);
        public static void ScreenUpdating(bool updateScreen) => PerformOperationWithRecovery.PerformOperation(() => ExcelApplication.Visible = updateScreen);
        public static void DisplayAlerts(bool displayAlerts) => PerformOperationWithRecovery.PerformOperation(() => ExcelApplication.DisplayAlerts = displayAlerts);

    }

}
