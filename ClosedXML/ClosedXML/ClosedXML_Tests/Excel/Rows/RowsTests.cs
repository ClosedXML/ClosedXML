using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class RowTests
    {

        [TestMethod]
        public void RowUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 2).SetValue("Test");
            ws.Cell(1, 3).SetValue("Test");

            var fromRow = ws.Row(1).RowUsed();
            Assert.AreEqual("B1:C1", fromRow.RangeAddress.ToStringRelative());

            var fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
            Assert.AreEqual("B1:C1", fromRange.RangeAddress.ToStringRelative());
        }


        [TestMethod]
        public void NoRowsUsed()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            Int32 count = 0;

            foreach (var row in ws.RowsUsed())
                count++;
            
            foreach (var row in ws.Range("A1:C3").RowsUsed())
                count++;

            Assert.AreEqual(0, count);
        }
    }
}
