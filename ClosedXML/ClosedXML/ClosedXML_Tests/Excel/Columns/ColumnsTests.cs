using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ColumnTests
    {

        [TestMethod]
        public void ColumnUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(2, 1).SetValue("Test");
            ws.Cell(3, 1).SetValue("Test");

            var fromColumn = ws.Column(1).ColumnUsed();
            Assert.AreEqual("A2:A3", fromColumn.RangeAddress.ToStringRelative());

            var fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
            Assert.AreEqual("A2:A3", fromRange.RangeAddress.ToStringRelative());
        }

    }
}
