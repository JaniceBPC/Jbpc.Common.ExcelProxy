using System;
using Jbpc.Common.Excel;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class BasicExcelAvailabilityTests
    {
        [Test]
        public void InstantiateExcelApplication()
        {
            Application excelApplication = null;

            try
            {
                excelApplication = ExcelOperations.ExcelApplication;
            }
            catch (Exception)
            {
                Assert.Fail("Unable to startup Excel Application");
            }

            Assert.NotNull(excelApplication, "Unable to startup Excel Application");
        }
        [Test]
        public void CreateWorkBook()
        {
            var workbook = ExcelOperations.NewWorkbook;

            Assert.NotNull(workbook);

            var worksheets = workbook.Worksheets;

            Assert.NotNull(worksheets);
        }
    }
}
