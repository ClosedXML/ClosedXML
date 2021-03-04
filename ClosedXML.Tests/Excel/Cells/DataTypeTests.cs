using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class DataTypeTests
    {
        [Test]
        public void ConvertNonNumericTextToNumberThrowsException()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                var c = ws.Cell("A1");
                c.Value = "ABC123";
                Assert.Throws<ArgumentException>(() => c.DataType = XLDataType.Number);
            }
        }
    }
}
