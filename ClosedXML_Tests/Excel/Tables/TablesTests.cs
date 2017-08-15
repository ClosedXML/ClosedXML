using ClosedXML.Attributes;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class TablesTests
    {
        public class TestObjectWithoutAttributes
        {
            public String Column1 { get; set; }
            public String Column2 { get; set; }
        }

        public class TestObjectWithAttributes
        {
            public int UnOrderedColumn { get; set; }

            [XLColumn(Header = "SecondColumn", Order = 1)]
            public String Column1 { get; set; }

            [XLColumn(Header = "FirstColumn", Order = 0)]
            public String Column2 { get; set; }

            [XLColumn(Header = "SomeFieldNotProperty", Order = 2)]
            public int MyField;
        }

        [Test]
        public void CanSaveTableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            var wb = new XLWorkbook();
            wb.AddWorksheet(dt);

            using (var ms = new MemoryStream())
                wb.SaveAs(ms, true);
        }

        [Test]
        public void CanSaveTableCreatedFromSingleRow()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title");
            ws.Range("A1").CreateTable();

            using (var ms = new MemoryStream())
                wb.SaveAs(ms, true);
        }

        [Test]
        public void CreatingATableFromHeadersPushCellsBelow()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title")
                .CellBelow().SetValue("X");
            ws.Range("A1").CreateTable();

            Assert.AreEqual(ws.Cell("A2").GetString(), String.Empty);
            Assert.AreEqual(ws.Cell("A3").GetString(), "X");
        }

        [Test]
        public void Inserting_Column_Sets_Header()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();
            table.InsertColumnsAfter(1);
            Assert.AreEqual("Column2", table.HeadersRow().LastCell().GetString());
        }

        [Test]
        public void SavingLoadingTableWithNewLineInHeader()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            string columnName = "Line1" + Environment.NewLine + "Line2";
            ws.FirstCell().SetValue(columnName)
                .CellBelow().SetValue("A");
            ws.RangeUsed().CreateTable();
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, true);
                var wb2 = new XLWorkbook(ms);
                IXLWorksheet ws2 = wb2.Worksheet(1);
                IXLTable table2 = ws2.Table(0);
                string fieldName = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName);
            }
        }

        [Test]
        public void SavingLoadingTableWithNewLineInHeader2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Test");

            var dt = new DataTable();
            string columnName = "Line1" + Environment.NewLine + "Line2";
            dt.Columns.Add(columnName);

            DataRow dr = dt.NewRow();
            dr[columnName] = "some text";
            dt.Rows.Add(dr);
            ws.Cell(1, 1).InsertTable(dt.AsEnumerable());

            IXLTable table1 = ws.Table(0);
            string fieldName1 = table1.Field(0).Name;
            Assert.AreEqual(columnName, fieldName1);

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms, true);
                var wb2 = new XLWorkbook(ms);
                IXLWorksheet ws2 = wb2.Worksheet(1);
                IXLTable table2 = ws2.Table(0);
                string fieldName2 = table2.Field(0).Name;
                Assert.AreEqual("Line1\nLine2", fieldName2);
            }
        }

        [Test]
        public void TableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());
        }

        [Test]
        public void TableCreatedFromEmptyListOfInt()
        {
            var l = new List<Int32>();

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(1, ws.Tables.First().ColumnCount());
        }

        [Test]
        public void TableCreatedFromEmptyListOfObject()
        {
            var l = new List<TestObjectWithoutAttributes>();

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(2, ws.Tables.First().ColumnCount());
        }

        [Test]
        public void TableCreatedFromListOfObjectWithPropertyAttributes()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.AreEqual(4, ws.Tables.First().ColumnCount());
            Assert.AreEqual("FirstColumn", ws.FirstCell().Value);
            Assert.AreEqual("SecondColumn", ws.FirstCell().CellRight().Value);
            Assert.AreEqual("SomeFieldNotProperty", ws.FirstCell().CellRight().CellRight().Value);
            Assert.AreEqual("UnOrderedColumn", ws.FirstCell().CellRight().CellRight().CellRight().Value);
        }

        [Test]
        public void TableInsertAboveFromData()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }

        [Test]
        public void TableInsertAboveFromRows()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 3;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsAbove(1).First();
            row.Field("Value").Value = 1;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }

        [Test]
        public void TableInsertBelowFromData()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = table.DataRange.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }

        [Test]
        public void TableInsertBelowFromRows()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Value");

            IXLTable table = ws.Range("A1:A2").CreateTable();
            table.SetShowTotalsRow()
                .Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;

            IXLTableRow row = table.DataRange.FirstRow();
            row.Field("Value").Value = 1;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 2;
            row = row.InsertRowsBelow(1).First();
            row.Field("Value").Value = 3;

            Assert.AreEqual(1, ws.Cell(2, 1).GetDouble());
            Assert.AreEqual(2, ws.Cell(3, 1).GetDouble());
            Assert.AreEqual(3, ws.Cell(4, 1).GetDouble());
        }

        [Test]
        public void TableShowHeader()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();

            Assert.AreEqual("Categories", table.Fields.First().Name);

            table.SetShowHeaderRow(false);

            Assert.AreEqual("Categories", table.Fields.First().Name);

            Assert.IsTrue(ws.Cell(1, 1).IsEmpty(true));
            Assert.AreEqual(null, table.HeadersRow());
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetString());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetString());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetString());
            Assert.AreEqual("C", table.DataRange.LastCell().GetString());

            table.SetShowHeaderRow();
            IXLRangeRow headerRow = table.HeadersRow();
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("Categories", headerRow.Cell(1).GetString());

            table.SetShowHeaderRow(false);

            ws.FirstCell().SetValue("x");

            table.SetShowHeaderRow();

            Assert.AreEqual("x", ws.FirstCell().GetString());
            Assert.AreEqual("Categories", ws.Cell("A2").GetString());
            Assert.AreNotEqual(null, headerRow);
            Assert.AreEqual("A", table.DataRange.FirstRow().Field("Categories").GetString());
            Assert.AreEqual("C", table.DataRange.LastRow().Field("Categories").GetString());
            Assert.AreEqual("A", table.DataRange.FirstCell().GetString());
            Assert.AreEqual("C", table.DataRange.LastCell().GetString());
        }

        [Test]
        public void ChangeFieldName()
        {
            XLWorkbook wb = new XLWorkbook();

            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").SetValue("FName")
                .CellBelow().SetValue("John");

            ws.Cell("B1").SetValue("LName")
                .CellBelow().SetValue("Doe");

            var tbl = ws.RangeUsed().CreateTable();
            var nameBefore = tbl.Field(tbl.Fields.Last().Index).Name;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastName";
            var nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;

            var cellValue = ws.Cell("B1").GetString();

            Assert.AreEqual("LName", nameBefore);
            Assert.AreEqual("LastName", nameAfter);
            Assert.AreEqual("LastName", cellValue);

            tbl.ShowHeaderRow = false;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastNameChanged";
            nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;
            Assert.AreEqual("LastNameChanged", nameAfter);

            tbl.SetShowHeaderRow(true);
            nameAfter = tbl.Cell("B1").Value.ToString();
            Assert.AreEqual("LastNameChanged", nameAfter);
        }
    }
}
