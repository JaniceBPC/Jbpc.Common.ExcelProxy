using System;
using System.Linq;
using System.Reflection;
using Jbpc.Common.Excel;
using Jbpc.Common.Excel.Proxies;
using NUnit.Framework;

namespace Jbpc.Common.UnitTests.Excel
{
    [TestFixture]
    public class ExcelTests
    {
        [Test]
        public void CreateWorkBook()
        {
            var workbook = ExcelOperationsProxy.NewWorkbook;

            Assert.NotNull(workbook);

            var worksheets = workbook.Worksheets;

            Assert.NotNull(worksheets);
        }
        [Test]
        public void OpenWorkbook()
        {
            var workbook = ExcelOperations.OpenReportWorkbook("BlankReport.xlsx");

            var worksheet = workbook.Worksheets.First();

            var range = worksheet.RangeForCell(1, 1);

            range.SetText("Hi Janice!");
        }
    }
}
