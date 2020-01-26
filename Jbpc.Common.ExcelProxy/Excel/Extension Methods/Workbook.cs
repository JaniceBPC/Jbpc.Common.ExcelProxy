using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel.ExtensionMethods
{
    public static class WorkbookExtensionMethods
    {
        public static Worksheet Worksheet(this Workbook workbook, string worksheetName)
        {
            try
            {
                return (Worksheet)workbook.Worksheets[worksheetName];
            }
            catch (Exception)
            {
                var wksMsg = string.Join($"{Environment.NewLine}\t", workbook.WorksheetsList().Select(x => x.Name));

                throw new ApplicationException($"{Environment.NewLine}Template worksheet=\"{worksheetName}\" not found in workbook=\"{workbook.Name}\"{Environment.NewLine}\t={wksMsg}");
            }
        }
        public static Worksheet CreateCopy(this Worksheet worksheet)
        {
            var workbook = worksheet.Parent as Workbook;

            var q = worksheet.Name;

            return workbook.CreateCopy(worksheet.Name);
        }
        public static Worksheet CreateCopyInWorkbook(this Worksheet worksheet, Workbook targetWorkbook)
        {
            worksheet.Copy(After: targetWorkbook.LastWorksheet());

            return targetWorkbook.LastWorksheet();
        }
        public static Worksheet CreateCopy(this Workbook workbook, string name)
        {
            var worksheet = workbook.Worksheet(name);

            worksheet.Copy(After: workbook.LastWorksheet());

            return workbook.LastWorksheet();
        }
        public static Worksheet LastWorksheet(this Workbook workbook) => (Worksheet) workbook.Worksheets[workbook.Worksheets.Count];
        public static List<Worksheet> WorksheetsList(this Workbook workbook)
        {
            var list = new List<Worksheet>();

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                list.Add(worksheet);
            }

            return list;
        }
        public static void DisplayGridlines(this Workbook workbook, bool onOff) => workbook.Application.ActiveWindow.DisplayGridlines = onOff;
        public static List<string> WorksheetNames(this Workbook workbook) => WorksheetsList(workbook).Select(x=> x.Name).ToList();
        public static Worksheet GetWorksheet(this Workbook workbook, string worksheetName)
        {
            System.Diagnostics.Debug.Assert(workbook!=null,$"Null workbook");

            try
            {
                if (worksheetName == "")
                {
                    worksheetName = workbook.WorksheetNames().First();
                }

                return (Worksheet) workbook.Worksheets[worksheetName];
            }
            catch
            {
                throw new ApplicationException($"Worksheet={worksheetName}, not found in workbook={workbook.Name}, worksheets={workbook.WorksheetNamesMsg()}"); 
            }
        }
        public static string WorksheetNamesMsg(this Workbook workbook) => string.Join(", ", workbook.WorksheetNames());
        public static  Range Range(this Workbook workbook, string worksheetName, int row, int col) => Worksheet(workbook, worksheetName).Cells[row, col] as Range;

    }
}
