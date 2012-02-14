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
        [TestMethod]
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
        
        [TestMethod]
        public void InsertingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [TestMethod]
        public void InsertingSheets2()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);
            wb.Worksheets.Add("Sheet3");

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [TestMethod]
        public void InsertingSheets3()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2", 1);
            wb.Worksheets.Add("Sheet1", 1);
            

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual("Sheet3", wb.Worksheet(3).Name);
        }

        [TestMethod]
        public void DeletingSheets1()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet3");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet1", 1);

            wb.Worksheet("Sheet3").Delete();

            Assert.AreEqual("Sheet1", wb.Worksheet(1).Name);
            Assert.AreEqual("Sheet2", wb.Worksheet(2).Name);
            Assert.AreEqual(2, wb.Worksheets.Count);
        }
        
    }
}
