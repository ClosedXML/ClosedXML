using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class TablesTests
    {

        [TestMethod]
        public void TableShowHeader()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            ws.RangeUsed().CreateTable().SetShowHeaderRow(false);

            var table = ws.Tables.First();

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox1.xlsx");

            Assert.IsTrue(ws.Cell(1,1).IsEmpty(true));
            Assert.AreEqual(null, table.HeadersRow());
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetString());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetString());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetString());
            Assert.AreEqual("C", table.DataRange.LastCell().GetString());

            table.SetShowHeaderRow();
            var headerRow = table.HeadersRow();
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("Categories", headerRow.Cell(1).GetString());

            

            table.SetShowHeaderRow(false);
            
            ws.FirstCell().SetValue("x");
            
            table.SetShowHeaderRow();
            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox2.xlsx");

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox3.xlsx");

            Assert.AreEqual("x", ws.FirstCell().GetString());
            Assert.AreEqual("Categories", ws.Cell("A2").GetString());
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetString());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetString());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetString());
            Assert.AreEqual("C", table.DataRange.LastCell().GetString());

        }


        [TestMethod]
        public void TableInsertBelowFromRows()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            var table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            var row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

        }

        [TestMethod]
        public void TableInsertBelowFromData()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            var table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            var row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());

            

        }

        [TestMethod]
        public void TableInsertAboveFromRows()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            var table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;
            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox1.xlsx");
            var row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;
            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox2.xlsx");

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");
        }

        [TestMethod]
        public void TableInsertAboveFromData()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            var table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            var row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

        }
    }
}
