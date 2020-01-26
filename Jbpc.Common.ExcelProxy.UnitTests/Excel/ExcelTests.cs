using System;
using System.Reflection;
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
    }
}
