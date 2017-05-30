using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class RangeUsedTests
    {
        [Test]
        public void CanGetNamedFromAnother()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").SetValue(1).AddToNamed("value1");
            
            Assert.AreEqual(1, wb.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, wb.Range("value1").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws1.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, ws1.Range("value1").FirstCell().GetValue<int>());

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

            Assert.AreEqual(1, wb.Cell("value2").GetValue<int>());
            Assert.AreEqual(1, wb.Range("value2").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws2.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, ws2.Range("value1").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws2.Cell("value2").GetValue<int>());
            Assert.AreEqual(1, ws2.Range("value2").FirstCell().GetValue<int>());
        }
    }
}