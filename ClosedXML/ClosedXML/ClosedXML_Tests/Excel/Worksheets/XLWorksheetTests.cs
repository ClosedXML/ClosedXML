using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;
using System.IO;

namespace ClosedXML_Tests
{
    [TestClass()]
    public class XLWorksheetTests
    {
        [TestMethod()]
        public void MergedRanges()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("A1:B2").Merge();
            ws.Range("C1:D3").Merge();
            ws.Range("D2:E2").Merge();
            
            Assert.AreEqual(2, ws.MergedRanges.Count);
            Assert.AreEqual("A1:B2", ws.MergedRanges.First().RangeAddress.ToStringRelative());
            Assert.AreEqual("D2:E2", ws.MergedRanges.Last().RangeAddress.ToStringRelative());
        }

    }
}
