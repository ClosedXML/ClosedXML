using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;
using System.IO;
using System.Drawing;

namespace ClosedXML_Tests
{
    [TestClass]
    public class MergedRangesTests
    {
        [TestMethod]
        public void LastCellFromMerge()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");
            ws.Range("B2:D4").Merge();

            var first = ws.FirstCellUsed(true).Address.ToStringRelative();
            var last = ws.LastCellUsed(true).Address.ToStringRelative();
            
            Assert.AreEqual("B2", first);
            Assert.AreEqual("D4", last);
        }


    }
}
