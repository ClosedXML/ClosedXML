using System;
using System.Data;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Tables
{
    [TestFixture]
    public class TableNameValidationTests
    {
        [Test]
        public void TestTableNameValidatorRules()
        {
            string message;
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(0);
                //Table names cannot be empty
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook(string.Empty, ws, out message));
                Assert.AreEqual("The table name '' is invalid", message);

                //Table names cannot be Whitespace
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook("   ", ws, out message));
                Assert.AreEqual("The table name '   ' is invalid", message);

                //Table names cannot be Null
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook(null, ws, out message));
                Assert.AreEqual("The table name '' is invalid", message);

                //Table names cannot start with number
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook("1Table", ws, out message));
                Assert.AreEqual("The table name '1Table' does not begin with a letter, an underscore or a backslash.",
                    message);

                //Strings cannot be longer then 255 charters
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook(
                    new string(Enumerable.Repeat('a', 256).ToArray()), ws, out message));
                Assert.AreEqual("The table name is more than 255 characters", message);

                //Table names cannot contain spaces
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook("Spaces in name", ws, out message));
                Assert.AreEqual("Table names cannot contain spaces", message);

                //Table names cannot be a cell address
                Assert.False(TableNameValidator.IsValidTableNameInWorkbook("R1C2", ws, out message));
                Assert.AreEqual("Table name cannot be a valid Cell Address 'R1C2'.", message);
            }
        }

        [Test]
        public void AssertCreatingTableWithSpaceInNameThrowsException()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                Assert.AreEqual("Table1", t1.Name);
                Assert.Throws<ArgumentException>(() => t1.Name = "Table name with spaces");
            }
        }

        [Test]
        public void AssertSettingExistingTableToSameNameDoesNotThrowException()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                Assert.AreEqual("Table1", t1.Name);
                Assert.DoesNotThrow(() => t1.Name = "TABLE1");
            }
        }

        [Test]
        public void AssertInsertingTableWithInvalidTableNamesThrowsException()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("Patient", typeof(string));
            dt.Rows.Add("David");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "May2019"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "A1"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "R1C2"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "r3c2"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "R2C33333"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "RC"));
                Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "RC"));
            }
        }

        [Test]
        public void TestTableMustBeUniqueAcrossTheWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                var t2 = ws1.Cell("G1").InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                Assert.AreEqual("Table1", t1.Name);
                Assert.AreEqual("Table2", t2.Name);
                var ex = Assert.Throws<ArgumentException>(() => t2.Name = "TABLE1");
                Assert.AreEqual("There is already a table named 'TABLE1'", ex?.Message);
            }
        }

        [Test]
        public void TestTableNameIsUniqueAcrossDefinedNames()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var ws2 = wb.AddWorksheet();

                //Create workbook scoped defined name
                wb.NamedRanges.Add("WorkbookScopedDefinedName", "Sheet1!A1:A10");
                ws2.NamedRanges.Add("WorksheetScopedDefinedName", "Sheet2!A1:A10");


                var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                var t2 = ws2.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
                Assert.AreEqual("Table1", t1.Name);
                Assert.AreEqual("Table2", t2.Name);

                var ex = Assert.Throws<ArgumentException>(() => t1.Name = "WorkbookScopedDefinedName");
                if (ex != null)
                    Assert.AreEqual(
                        "Table name must be unique across all named ranges 'WorkbookScopedDefinedName'.",
                        ex.Message);

                ex = Assert.Throws<ArgumentException>(() => t2.Name = "WorksheetScopedDefinedName");
                if (ex != null)
                    Assert.AreEqual(
                        "Table name must be unique across all named ranges 'WorksheetScopedDefinedName'.",
                        ex.Message);
            }
        }
    }
}
