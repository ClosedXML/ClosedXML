using System.Collections.Generic;
using System.Data;
using System.IO;
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
        public void Inserting_Column_Sets_Header()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            var table = ws.RangeUsed().CreateTable();
            table.InsertColumnsAfter(1);
            Assert.AreEqual("Column2", table.HeadersRow().LastCell().GetString());
        }

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

            //wb.SaveAs(@"D:\Excel Files\ForTesting\Sandbox.xlsx");

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

        [TestMethod]
        public void CreatingATableFromHeadersPushCellsBelow()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title")
                .CellBelow().SetValue("X");
            ws.Range("A1").CreateTable();

            Assert.AreEqual(ws.Cell("A2").GetString(), String.Empty);
            Assert.AreEqual(ws.Cell("A3").GetString(), "X");
        }


        [TestMethod]
        public void CanSaveTableCreatedFromSingleRow()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title");
            ws.Range("A1").CreateTable();

            using (var ms = new MemoryStream())
                wb.SaveAs(ms);

        }

        [TestMethod]
        public void CanSaveTableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            var wb = new XLWorkbook();
            wb.AddWorksheet(dt);

            using (var ms = new MemoryStream())
                wb.SaveAs(ms);

        }

        [TestMethod]
        public void TableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());
        }

        [TestMethod]
        public void TableCreatedFromEmptyListOfInt()
        {
            var l = new List<Int32>();

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(1, ws.Tables.First().ColumnCount());

        }

        public class TestObject
        {
            public String Column1 { get; set; }
            public String Column2 { get; set; }
        }
        [TestMethod]
        public void TableCreatedFromEmptyListOfObject()
        {
            var l = new List<TestObject>();

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());

        }

        [TestMethod]
        public void SavingLoadingTableWithNewLineInHeader()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var columnName = "Line1" + Environment.NewLine + "Line2";
            ws.FirstCell().SetValue(columnName)
                .CellBelow().SetValue("A");
            ws.RangeUsed().CreateTable();
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                var wb2 = new XLWorkbook(ms);
                var ws2 = wb2.Worksheet(1);
                var table2 = ws2.Table(0);
                var fieldName = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName);
            }

        }

        [TestMethod]
        public void SavingLoadingTableWithNewLineInHeader2()
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Test");

            DataTable dt = new DataTable();
            var columnName = "Line1" + Environment.NewLine + "Line2";
            dt.Columns.Add(columnName);

            DataRow dr = dt.NewRow();
            dr[columnName] = "some text";
            dt.Rows.Add(dr);
            ws.Cell(1, 1).InsertTable(dt.AsEnumerable());

            var table1 = ws.Table(0);
            var fieldName1 = table1.Field(0).Name;
            Assert.AreEqual(columnName, fieldName1);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                var wb2 = new XLWorkbook(ms);
                var ws2 = wb2.Worksheet(1);
                var table2 = ws2.Table(0);
                var fieldName2 = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName2);
            }

        }
    }
}
