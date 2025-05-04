using ClosedXML.Attributes;
using ClosedXML.Excel;
using ClosedXML.Excel.Exceptions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class TablesTests
    {
        public class TestObjectWithoutAttributes
        {
            public string Column1 { get; set; }
            public string Column2 { get; set; }
        }

        public class TestObjectWithAttributes
        {
            public int UnOrderedColumn { get; set; }

            [XLColumn(Header = "SecondColumn", Order = 1)]
            public string Column1 { get; set; }

            [XLColumn(Header = "FirstColumn", Order = 0)]
            public string Column2 { get; set; }

            [XLColumn(Header = "SomeFieldNotProperty", Order = 2)]
            public int MyField;
        }

        [Test]
        public void CanSaveTableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            using var wb = new XLWorkbook();
            wb.AddWorksheet(dt);

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        }

        [Test]
        public void PreventAddingOfEmptyDataTable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var dt = new DataTable();
            var table = ws.FirstCell().InsertTable(dt);

            Assert.That(table, Is.Null);
        }

        [Test]
        public void CanSaveTableCreatedFromSingleRow()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title");
            ws.Range("A1").CreateTable();

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        }

        [Test]
        public void CreatingATableFromHeadersPushCellsBelow()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Title")
                .CellBelow().SetValue("X");
            ws.Range("A1").CreateTable();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(Blank.Value));
                Assert.That(ws.Cell("A3").GetText(), Is.EqualTo("X"));
            });
        }

        [Test]
        public void Inserting_Column_Sets_Header()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();
            table.InsertColumnsAfter(1);
            Assert.That(table.HeadersRow().LastCell().GetText(), Is.EqualTo("Column2"));
        }

        [Test]
        public void DataRange_returns_null_if_empty()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();

            ws.Rows("2:4").Delete();

            Assert.That(table.DataRange, Is.Null);
        }

        [Test]
        public void SavingLoadingTableWithNewLineInHeader()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            string columnName = "Line1" + Environment.NewLine + "Line2";
            ws.FirstCell().SetValue(columnName)
                .CellBelow().SetValue("A");
            ws.RangeUsed().CreateTable();
            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
            var wb2 = new XLWorkbook(ms);
            IXLWorksheet ws2 = wb2.Worksheet(1);
            IXLTable table2 = ws2.Table(0);
            string fieldName = table2.Field(0).Name;
            Assert.That(fieldName, Is.EqualTo("Line1\nLine2"));
        }

        [Test]
        public void SavingLoadingTableWithNewLineInHeader2()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Test");

            var dt = new DataTable();
            string columnName = "Line1" + Environment.NewLine + "Line2";
            dt.Columns.Add(columnName);

            DataRow dr = dt.NewRow();
            dr[columnName] = "some text";
            dt.Rows.Add(dr);
            ws.Cell(1, 1).InsertTable(dt);

            IXLTable table1 = ws.Table(0);
            string fieldName1 = table1.Field(0).Name;
            Assert.That(fieldName1, Is.EqualTo(columnName));

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
            var wb2 = new XLWorkbook(ms);
            IXLWorksheet ws2 = wb2.Worksheet(1);
            IXLTable table2 = ws2.Table(0);
            string fieldName2 = table2.Field(0).Name;
            Assert.That(fieldName2, Is.EqualTo("Line1\nLine2"));
        }

        [Test]
        public void TableCreatedFromEmptyDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt);
            Assert.That(ws.Tables.First().ColumnCount(), Is.EqualTo(2));
        }

        [Test]
        public void TableCreatedFromEmptyListOfInt()
        {
            var l = new List<int>();

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.That(ws.Tables.First().ColumnCount(), Is.EqualTo(1));
        }

        [Test]
        public void TableCreatedFromEmptyListOfObject()
        {
            var l = new List<TestObjectWithoutAttributes>();

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.That(ws.Tables.First().ColumnCount(), Is.EqualTo(2));
        }

        [Test]
        public void TableCreatedFromListOfObjectWithPropertyAttributes()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.Multiple(() =>
            {
                Assert.That(ws.Tables.First().ColumnCount(), Is.EqualTo(4));
                Assert.That(ws.FirstCell().Value, Is.EqualTo("FirstColumn"));
                Assert.That(ws.FirstCell().CellRight().Value, Is.EqualTo("SecondColumn"));
                Assert.That(ws.FirstCell().CellRight().CellRight().Value, Is.EqualTo("SomeFieldNotProperty"));
                Assert.That(ws.FirstCell().CellRight().CellRight().CellRight().Value, Is.EqualTo("UnOrderedColumn"));
            });
        }

        [Test]
        public void EmptyTableCreatedFromListOfObjectWithPropertyAttributes()
        {
            var l = new List<TestObjectWithAttributes>();

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.Multiple(() =>
            {
                Assert.That(ws.Tables.First().ColumnCount(), Is.EqualTo(4));
                Assert.That(ws.FirstCell().Value, Is.EqualTo("FirstColumn"));
                Assert.That(ws.FirstCell().CellRight().Value, Is.EqualTo("SecondColumn"));
                Assert.That(ws.FirstCell().CellRight().CellRight().Value, Is.EqualTo("SomeFieldNotProperty"));
                Assert.That(ws.FirstCell().CellRight().CellRight().CellRight().Value, Is.EqualTo("UnOrderedColumn"));
            });
        }

        [Test]
        public void TableInsertAboveFromData()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(2, 1).GetDouble(), Is.EqualTo(1));
                Assert.That(ws.Cell(3, 1).GetDouble(), Is.EqualTo(2));
                Assert.That(ws.Cell(4, 1).GetDouble(), Is.EqualTo(3));
            });
        }

        [Test]
        public void TableInsertAboveFromRows()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(2, 1).GetDouble(), Is.EqualTo(1));
                Assert.That(ws.Cell(3, 1).GetDouble(), Is.EqualTo(2));
                Assert.That(ws.Cell(4, 1).GetDouble(), Is.EqualTo(3));
            });
        }

        [Test]
        public void TableInsertBelowFromData()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(2, 1).GetDouble(), Is.EqualTo(1));
                Assert.That(ws.Cell(3, 1).GetDouble(), Is.EqualTo(2));
                Assert.That(ws.Cell(4, 1).GetDouble(), Is.EqualTo(3));
            });
        }

        [Test]
        public void TableInsertBelowFromRows()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(2, 1).GetDouble(), Is.EqualTo(1));
                Assert.That(ws.Cell(3, 1).GetDouble(), Is.EqualTo(2));
                Assert.That(ws.Cell(4, 1).GetDouble(), Is.EqualTo(3));
            });
        }

        [Test]
        public void TableShowHeader()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B")
                .CellBelow().SetValue("C");

            IXLTable table = ws.RangeUsed().CreateTable();

            Assert.That(table.Fields.First().Name, Is.EqualTo("Categories"));

            table.SetShowHeaderRow(false);

            Assert.Multiple(() =>
            {
                Assert.That(table.Fields.First().Name, Is.EqualTo("Categories"));

                Assert.That(ws.Cell(1, 1).IsEmpty(XLCellsUsedOptions.All), Is.True);
                Assert.That(table.HeadersRow(), Is.Null);
                Assert.That(table.DataRange.FirstRow().Field("Categories").GetText(), Is.EqualTo("A"));
                Assert.That(table.DataRange.LastRow().Field("Categories").GetText(), Is.EqualTo("C"));
                Assert.That(table.DataRange.FirstCell().GetText(), Is.EqualTo("A"));
                Assert.That(table.DataRange.LastCell().GetText(), Is.EqualTo("C"));
            });

            table.SetShowHeaderRow();
            IXLRangeRow headerRow = table.HeadersRow();
            Assert.That(headerRow, Is.Not.Null);
            Assert.That(headerRow.Cell(1).GetText(), Is.EqualTo("Categories"));

            table.SetShowHeaderRow(false);

            ws.FirstCell().SetValue("x");

            table.SetShowHeaderRow();

            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().GetText(), Is.EqualTo("x"));
                Assert.That(ws.Cell("A2").GetText(), Is.EqualTo("Categories"));
            });
            Assert.That(headerRow, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(table.DataRange.FirstRow().Field("Categories").GetText(), Is.EqualTo("A"));
                Assert.That(table.DataRange.LastRow().Field("Categories").GetText(), Is.EqualTo("C"));
                Assert.That(table.DataRange.FirstCell().GetText(), Is.EqualTo("A"));
                Assert.That(table.DataRange.LastCell().GetText(), Is.EqualTo("C"));
            });
        }

        [TestCase("Amount")]
        [TestCase("AMOUNT")]
        [TestCase("amount")]
        public void FieldNames_of_XLTable_are_case_insensitive(string fieldName)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = ws.Cell("A1").InsertTable(new[] { new { Amount = 1 } });

            var expectedField = table.Field(0);
            Assert.That(table.Field(fieldName), Is.SameAs(expectedField));
        }

        [Test]
        public void ChangeFieldName()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").SetValue("FName")
                .CellBelow().SetValue("John");

            ws.Cell("B1").SetValue("LName")
                .CellBelow().SetValue("Doe");

            var tbl = ws.RangeUsed().CreateTable();
            var nameBefore = tbl.Field(tbl.Fields.Last().Index).Name;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastName";
            var nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;

            var cellValue = ws.Cell("B1").GetText();

            Assert.Multiple(() =>
            {
                Assert.That(nameBefore, Is.EqualTo("LName"));
                Assert.That(nameAfter, Is.EqualTo("LastName"));
                Assert.That(cellValue, Is.EqualTo("LastName"));
            });

            tbl.ShowHeaderRow = false;
            tbl.Field(tbl.Fields.Last().Index).Name = "LastNameChanged";
            nameAfter = tbl.Field(tbl.Fields.Last().Index).Name;
            Assert.That(nameAfter, Is.EqualTo("LastNameChanged"));

            tbl.SetShowHeaderRow(true);
            nameAfter = (string)tbl.Cell("B1").Value;
            Assert.That(nameAfter, Is.EqualTo("LastNameChanged"));

            var field = tbl.Field("LastNameChanged");
            Assert.That(field.Name, Is.EqualTo("LastNameChanged"));

            tbl.Cell(1, 1).Value = "FirstName";
            Assert.That(tbl.Field(0).Name, Is.EqualTo("FirstName"));
        }

        [Test]
        public void CanDeleteTableColumn()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var table = ws.FirstCell().InsertTable(l);

            table.Column("C").Delete();

            Assert.Multiple(() =>
            {
                Assert.That(table.Fields.Count(), Is.EqualTo(3));

                Assert.That(table.Fields.First().Name, Is.EqualTo("FirstColumn"));
                Assert.That(table.Fields.First().Index, Is.EqualTo(0));

                Assert.That(table.Fields.Last().Name, Is.EqualTo("UnOrderedColumn"));
                Assert.That(table.Fields.Last().Index, Is.EqualTo(2));
            });
        }

        [Test]
        public void TestFieldCellTypes()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var table = ws.Cell("B2").InsertTable(l);

            Assert.Multiple(() =>
            {
                Assert.That(table.Fields.Count(), Is.EqualTo(4));
                Assert.That(table.Field(0).HeaderCell.Address.ToString(), Is.EqualTo("B2"));
            });
            Assert.Multiple(() =>
            {
                Assert.That(table.Field(1).HeaderCell.Address.ToString(), Is.EqualTo("C2"));
                Assert.That(table.Field(2).HeaderCell.Address.ToString(), Is.EqualTo("D2"));
                Assert.That(table.Field(3).HeaderCell.Address.ToString(), Is.EqualTo("E2"));

                Assert.That(table.Field(0).TotalsCell, Is.Null);
                Assert.That(table.Field(1).TotalsCell, Is.Null);
                Assert.That(table.Field(2).TotalsCell, Is.Null);
                Assert.That(table.Field(3).TotalsCell, Is.Null);
            });

            table.SetShowTotalsRow();

            Assert.That(table.Field(0).TotalsCell.Address.ToString(), Is.EqualTo("B5"));
            Assert.That(table.Field(1).TotalsCell.Address.ToString(), Is.EqualTo("C5"));
            Assert.That(table.Field(2).TotalsCell.Address.ToString(), Is.EqualTo("D5"));
            Assert.That(table.Field(3).TotalsCell.Address.ToString(), Is.EqualTo("E5"));

            var field = table.Fields.Last();

            Assert.Multiple(() =>
            {
                Assert.That(field.Column.RangeAddress.ToString(), Is.EqualTo("E2:E5"));
                Assert.That(field.DataCells.First().Address.ToString(), Is.EqualTo("E3"));
                Assert.That(field.DataCells.Last().Address.ToString(), Is.EqualTo("E4"));
            });
        }

        [Test]
        public void CanDeleteTable()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().InsertTable(l);
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var table = ws.Tables.First();

                ws.Tables.Remove(table.Name);
                Assert.That(ws.Tables.Count(), Is.EqualTo(0));
                wb.Save();
            }
        }

        [Test]
        public void TableNameCannotBeValidCellName()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("Patient", typeof(string));
            dt.Rows.Add("David");

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "May2019"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "A1"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "R1C2"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "r3c2"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "R2C33333"));
            Assert.Throws<InvalidOperationException>(() => ws.Cell(1, 1).InsertTable(dt, "RC"));
        }

        [Test]
        public void TableNameSetWhenAddingWorksheetWithDataTable()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("Patient", typeof(string));
            dt.Rows.Add("David");

            using (var wb = new XLWorkbook())
            {
                // Generated table name is used and should not be an issue
                Assert.DoesNotThrow(() => wb.AddWorksheet(dt, "t1"));
            }

            using (var wb = new XLWorkbook())
            {
                // Should pass because t1 is a valid sheet name, and is not used for the tableName
                Assert.DoesNotThrow(() => wb.AddWorksheet(dt, "t1", "table1"));

                Assert.Multiple(() =>
                {
                    Assert.That(wb.Worksheets, Has.Count.EqualTo(1));
                    Assert.That(wb.Worksheet(1).Tables.Count(), Is.EqualTo(1));
                });
            }

        }

        [Test]
        public void CanDeleteTableField()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var table = ws.Cell("B2").InsertTable(l);

            Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:E4"));

            table.Field("SomeFieldNotProperty").Delete();

            Assert.Multiple(() =>
            {
                Assert.That(table.Fields.Count(), Is.EqualTo(3));

                Assert.That(table.Fields.First().Name, Is.EqualTo("FirstColumn"));
                Assert.That(table.Fields.First().Index, Is.EqualTo(0));

                Assert.That(table.Fields.Last().Name, Is.EqualTo("UnOrderedColumn"));
                Assert.That(table.Fields.Last().Index, Is.EqualTo(2));

                Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:D4"));
            });
        }

        [Test]
        public void CanDeleteTableRows()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 },
                new TestObjectWithAttributes() { Column1 = "e", Column2 = "f", MyField = 6, UnOrderedColumn = 555 },
                new TestObjectWithAttributes() { Column1 = "g", Column2 = "h", MyField = 7, UnOrderedColumn = 333 }
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var table = ws.Cell("B2").InsertTable(l);

            Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:E6"));

            table.DataRange.Rows(3, 4).Delete();

            Assert.Multiple(() =>
            {
                Assert.That(table.DataRange.Rows().Count(), Is.EqualTo(2));

                Assert.That(table.DataRange.FirstCell().Value, Is.EqualTo("b"));
                Assert.That(table.DataRange.LastCell().Value, Is.EqualTo(777));

                Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:E4"));
            });
        }

        [Test]
        public void OverlappingTablesThrowsException()
        {
            var dt = new DataTable("sheet1");
            dt.Columns.Add("col1", typeof(string));
            dt.Columns.Add("col2", typeof(double));

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(dt, true);
            Assert.Throws<InvalidOperationException>(() => ws.FirstCell().CellRight().InsertTable(dt, true));
        }

        [Test]
        public void OverwritingTableHeaders()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = ws.Cell("A1").InsertTable(new object[]
            {
                ("Header 1", "Header 2"),
                (1, 2)
            }, true);

            // Overwrite the headers of the table with non-string values
            ws.Cell("A1").InsertData(new object[]
            {
                (XLError.IncompatibleValue, 7)
            });

            Assert.Multiple(() =>
            {
                // The non-string data inserted to headers were converted to strings and used as a field names.
                Assert.That(table.Field(0).Name, Is.EqualTo("#VALUE!"));
                Assert.That(ws.Cell("A1").Value, Is.EqualTo("#VALUE!"));
                Assert.That(table.Field(1).Name, Is.EqualTo("7"));
                Assert.That(ws.Cell("B1").Value, Is.EqualTo("7"));
            });
        }

        [Test]
        public void OverwritingTableTotalsRow()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable.Range(1, 10)
                .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new string('a', i)
                    });

            var table = ws.FirstCell().InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            var data2 = Enumerable.Range(1, 20)
                .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new string('b', i),
                        Int = 64 + i
                    });

            ws.FirstCell().CellBelow().InsertData(data2);

            table.Fields.ForEach(f => Assert.That(f.TotalsRowFunction, Is.EqualTo(XLTotalsRowFunction.None)));

            Assert.Multiple(() =>
            {
                Assert.That(table.Field(0).TotalsRowLabel, Is.EqualTo("11"));
                Assert.That(table.Field(1).TotalsRowLabel, Is.EqualTo("K"));
                Assert.That(table.Field(2).TotalsRowLabel, Is.EqualTo("bbbbbbbbbbb"));
            });
        }

        [Test]
        public void TableRenameTests()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var table1 = ws.FirstCell().InsertTable(l);
            var table2 = ws.Cell("A10").InsertTable(l);

            Assert.Multiple(() =>
            {
                Assert.That(table1.Name, Is.EqualTo("Table1"));
                Assert.That(table2.Name, Is.EqualTo("Table2"));
            });

            table1.Name = "table1";
            Assert.That(table1.Name, Is.EqualTo("table1"));

            table1.Name = "_table1";
            Assert.That(table1.Name, Is.EqualTo("_table1"));

            table1.Name = "\\table1";
            Assert.That(table1.Name, Is.EqualTo("\\table1"));

            Assert.Throws<ArgumentException>(() => table1.Name = "");
            Assert.Throws<ArgumentException>(() => table1.Name = "R");
            Assert.Throws<ArgumentException>(() => table1.Name = "C");
            Assert.Throws<ArgumentException>(() => table1.Name = "r");
            Assert.Throws<ArgumentException>(() => table1.Name = "c");

            Assert.Throws<ArgumentException>(() => table1.Name = "123");
            Assert.Throws<ArgumentException>(() => table1.Name = new string('A', 256));

            Assert.Throws<ArgumentException>(() => table1.Name = "Table2");
            Assert.Throws<ArgumentException>(() => table1.Name = "TABLE2");
        }

        [Test]
        public void CanResizeTable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable.Range(1, 10)
                .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new string('a', i)
                    });

            var table = ws.FirstCell().InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            var data2 = Enumerable.Range(1, 10)
                .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new string('b', i),
                        Integer = 64 + i
                    });

            ws.FirstCell().CellBelow().InsertData(data2);
            table.Resize(table.FirstCell().Address, table.AsRange().LastCell().CellRight().Address);

            Assert.Multiple(() =>
            {
                Assert.That(table.Fields.Count(), Is.EqualTo(4));

                Assert.That(table.Field(3).Name, Is.EqualTo("Column4"));
            });

            ws.Cell("D1").Value = "Integer";
            Assert.That(table.Field(3).Name, Is.EqualTo("Integer"));
        }

        [Test]
        public void TableAsDynamicEnumerable()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var table = ws.FirstCell().InsertTable(l);

            foreach (var d in table.AsDynamicEnumerable())
            {
                Assert.DoesNotThrow(() =>
                {
                    object value;
                    value = d.FirstColumn;
                    value = d.SecondColumn;
                    value = d.UnOrderedColumn;
                    value = d.SomeFieldNotProperty;
                });
            }
        }

        [Test]
        public void TableAsDotNetDataTable()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            var table = ws.FirstCell().InsertTable(l).AsNativeDataTable();

            Assert.That(table.Columns, Has.Count.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(table.Columns[0].ColumnName, Is.EqualTo("FirstColumn"));
                Assert.That(table.Columns[1].ColumnName, Is.EqualTo("SecondColumn"));
                Assert.That(table.Columns[2].ColumnName, Is.EqualTo("SomeFieldNotProperty"));
                Assert.That(table.Columns[3].ColumnName, Is.EqualTo("UnOrderedColumn"));

                Assert.That(table.Columns[0].DataType, Is.EqualTo(typeof(string)));
                Assert.That(table.Columns[1].DataType, Is.EqualTo(typeof(string)));
                Assert.That(table.Columns[2].DataType, Is.EqualTo(typeof(double)));
                Assert.That(table.Columns[3].DataType, Is.EqualTo(typeof(double)));
            });

            var dr = table.Rows[0];
            Assert.Multiple(() =>
            {
                Assert.That(dr["FirstColumn"], Is.EqualTo("b"));
                Assert.That(dr["SecondColumn"], Is.EqualTo("a"));
                Assert.That(dr["SomeFieldNotProperty"], Is.EqualTo(4));
                Assert.That(dr["UnOrderedColumn"], Is.EqualTo(999));
            });

            dr = table.Rows[1];
            Assert.Multiple(() =>
            {
                Assert.That(dr["FirstColumn"], Is.EqualTo("d"));
                Assert.That(dr["SecondColumn"], Is.EqualTo("c"));
                Assert.That(dr["SomeFieldNotProperty"], Is.EqualTo(5));
                Assert.That(dr["UnOrderedColumn"], Is.EqualTo(777));
            });
        }

        [Test]
        public void TestTableCellTypes()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var data1 = Enumerable.Range(1, 10)
                .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new string('a', i)
                    });

            var table = ws.FirstCell().InsertTable(data1, true)
                .SetShowHeaderRow()
                .SetShowTotalsRow();
            table.Fields.First().TotalsRowFunction = XLTotalsRowFunction.Sum;

            Assert.Multiple(() =>
            {
                Assert.That(table.HeadersRow().Cell(1).TableCellType(), Is.EqualTo(XLTableCellType.Header));
                Assert.That(table.HeadersRow().Cell(1).CellBelow().TableCellType(), Is.EqualTo(XLTableCellType.Data));
                Assert.That(table.TotalsRow().Cell(1).TableCellType(), Is.EqualTo(XLTableCellType.Total));
                Assert.That(ws.Cell("Z100").TableCellType(), Is.EqualTo(XLTableCellType.None));
            });
        }

        [Test]
        public void TotalsFunctionsOfHeadersWithWeirdCharacters()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l, false);

            // Give the headings weird names (i.e. spaces, hashes, single quotes
            ws.Cell("A1").Value = "ABCD    ";
            ws.Cell("B1").Value = "   #BCD";
            ws.Cell("C1").Value = "   as'df   ";
            ws.Cell("D1").Value = "Normal";

            var table = ws.RangeUsed().CreateTable();
            Assert.IsNotNull(table);

            table.ShowTotalsRow = true;
            table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Count;
            table.Field(1).TotalsRowFunction = XLTotalsRowFunction.Count;
            table.Field(2).TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field(3).TotalsRowFunction = XLTotalsRowFunction.Sum;

            Assert.Multiple(() =>
            {
                Assert.That(table.Field(0).TotalsRowFormulaA1, Is.EqualTo("SUBTOTAL(103,Table1[[ABCD    ]])"));
                Assert.That(table.Field(1).TotalsRowFormulaA1, Is.EqualTo("SUBTOTAL(103,Table1[[   '#BCD]])"));
                Assert.That(table.Field(2).TotalsRowFormulaA1, Is.EqualTo("SUBTOTAL(109,Table1[[   as''df   ]])"));
                Assert.That(table.Field(3).TotalsRowFormulaA1, Is.EqualTo("SUBTOTAL(109,[Normal])"));
            });
        }

        [Test]
        public void CannotCreateDuplicateTablesOverSameRange()
        {
            var l = new List<TestObjectWithAttributes>()
            {
                new TestObjectWithAttributes() { Column1 = "a", Column2 = "b", MyField = 4, UnOrderedColumn = 999 },
                new TestObjectWithAttributes() { Column1 = "c", Column2 = "d", MyField = 5, UnOrderedColumn = 777 }
            };

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().InsertTable(l);
            Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().CreateTable());
        }

        [Test]
        public void CannotCreateTableOverExistingAutoFilter()
        {
            using var wb = new XLWorkbook();

            var data = Enumerable.Range(1, 10).Select(i => new
            {
                Index = i,
                String = $"String {i}"
            });

            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertTable(data, createTable: false);
            ws.RangeUsed().SetAutoFilter().Column(1).AddFilter(5);

            Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().CreateTable());
        }

        [Test]
        public void CopyTableSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");

            var table = ws1.Range("A1:C2").AsTable();

            TestDelegate action = () => table.CopyTo(ws1);

            Assert.Throws(typeof(InvalidOperationException), action);
        }

        [Test]
        public void CanInsertDateTimeOffset()
        {
            var now = DateTimeOffset.Now;

            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            ws1.FirstCell().InsertTable(new[] { new { TimeStamp = now } });

            // C# Supports 7 digits milliseconds, but excel only 3
            const string format = "yyyy-MM-dd HH:mm:ss.fff";

            var actual = ws1.Cell("A2").GetDateTime().ToString(format);
            var expected = now.DateTime.ToString(format);
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void CopyDetachedTableDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").Value = "Custom column 1";
            ws1.Cell("B1").Value = "Custom column 2";
            ws1.Cell("C1").Value = "Custom column 3";
            ws1.Cell("A2").Value = "Value 1";
            ws1.Cell("B2").Value = 123.45;
            ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
            var original = ws1.Range("A1:C2").AsTable("Detached table");
            var ws2 = wb.Worksheets.Add("Sheet2");

            var copy = original.CopyTo(ws2);

            Assert.Multiple(() =>
            {
                Assert.That(ws1.Tables.Count(), Is.EqualTo(0)); // We did not add it
                Assert.That(ws2.Tables.Count(), Is.EqualTo(1));
            });

            AssertTablesAreEqual(original, copy);

            Assert.Multiple(() =>
            {
                Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!A1:C2"));
                Assert.That(ws2.Cell("A1").Value, Is.EqualTo("Custom column 1"));
                Assert.That(ws2.Cell("B1").Value, Is.EqualTo("Custom column 2"));
                Assert.That(ws2.Cell("C1").Value, Is.EqualTo("Custom column 3"));
                Assert.That(ws2.Cell("A2").Value, Is.EqualTo("Value 1"));
                Assert.That((double)ws2.Cell("B2").Value, Is.EqualTo(123.45).Within(XLHelper.Epsilon));
                Assert.That(ws2.Cell("C2").Value, Is.EqualTo(new DateTime(2018, 5, 10)));
            });
        }

        [Test]
        public void CopyTableDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").Value = "Custom column 1";
            ws1.Cell("B1").Value = "Custom column 2";
            ws1.Cell("C1").Value = "Custom column 3";
            ws1.Cell("A2").Value = "Value 1";
            ws1.Cell("B2").Value = 123.45;
            ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
            var original = ws1.Range("A1:C2").AsTable("Attached table");
            ws1.Tables.Add(original);
            var ws2 = wb.Worksheets.Add("Sheet2");

            original.CopyTo(ws2);

            Assert.Multiple(() =>
            {
                Assert.That(ws1.Tables.Count(), Is.EqualTo(1));
                Assert.That(ws2.Tables.Count(), Is.EqualTo(1));
            });

            var copy = ws2.Tables.First();

            AssertTablesAreEqual(original, copy);

            Assert.Multiple(() =>
            {
                Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!A1:C2"));
                Assert.That(ws2.Cell("A1").Value, Is.EqualTo("Custom column 1"));
                Assert.That(ws2.Cell("B1").Value, Is.EqualTo("Custom column 2"));
                Assert.That(ws2.Cell("C1").Value, Is.EqualTo("Custom column 3"));
                Assert.That(ws2.Cell("A2").Value, Is.EqualTo("Value 1"));
                Assert.That((double)ws2.Cell("B2").Value, Is.EqualTo(123.45).Within(XLHelper.Epsilon));
                Assert.That(ws2.Cell("C2").Value, Is.EqualTo(new DateTime(2018, 5, 10)));
            });
        }

        [Test]
        public void NewTableHasNullRelId()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = "Custom column 1";
                ws.Cell("B1").Value = "Custom column 2";
                ws.Cell("C1").Value = "Custom column 3";
                ws.Cell("A2").Value = "Value 1";
                ws.Cell("B2").Value = 123.45;
                ws.Cell("C2").Value = new DateTime(2018, 5, 10);
                var original = ws.Range("A1:C2").CreateTable("Attached table");

                Assert.That(ws.Tables.Count(), Is.EqualTo(1));
                Assert.That((original as XLTable).RelId, Is.Null);

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.Add("Sheet2");
                var original = wb.Worksheets.First().Tables.First();

                Assert.IsNotNull((original as XLTable).RelId);

                var copy = original.CopyTo(ws);

                Assert.That(ws.Tables.Count(), Is.EqualTo(1));
                Assert.That((copy as XLTable).RelId, Is.Null);

                AssertTablesAreEqual(original, copy);

                Assert.Multiple(() =>
                {
                    Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!A1:C2"));
                    Assert.That(ws.Cell("A1").Value, Is.EqualTo("Custom column 1"));
                    Assert.That(ws.Cell("B1").Value, Is.EqualTo("Custom column 2"));
                    Assert.That(ws.Cell("C1").Value, Is.EqualTo("Custom column 3"));
                    Assert.That(ws.Cell("A2").Value, Is.EqualTo("Value 1"));
                    Assert.That((double)ws.Cell("B2").Value, Is.EqualTo(123.45).Within(XLHelper.Epsilon));
                    Assert.That(ws.Cell("C2").Value, Is.EqualTo(new DateTime(2018, 5, 10)));
                });
            }
        }

        [Test]
        public void CopyTableWithoutData()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").Value = "Custom column 1";
            ws1.Cell("B1").Value = "Custom column 2";
            ws1.Cell("C1").Value = "Custom column 3";
            ws1.Cell("A2").Value = "Value 1";
            ws1.Cell("B2").Value = 123.45;
            ws1.Cell("C2").Value = new DateTime(2018, 5, 10);
            var original = ws1.Range("A1:C2").AsTable("Attached table");
            ws1.Tables.Add(original);
            var ws2 = wb.Worksheets.Add("Sheet2") as XLWorksheet;

            var copy = (original as XLTable).CopyTo(ws2, false);

            AssertTablesAreEqual(original, copy);

            Assert.Multiple(() =>
            {
                Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!A1:C2"));
                Assert.That(ws2.Cell("A1").Value, Is.EqualTo("Custom column 1"));
                Assert.That(ws2.Cell("B1").Value, Is.EqualTo("Custom column 2"));
                Assert.That(ws2.Cell("C1").Value, Is.EqualTo("Custom column 3"));
                Assert.That(ws2.Cell("A2").Value, Is.EqualTo(Blank.Value));
                Assert.That(ws2.Cell("B2").Value, Is.EqualTo(Blank.Value));
                Assert.That(ws2.Cell("C2").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void SavingTableWithNullDataRangeThrowsException()
        {
            using var ms = new MemoryStream();
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var data = Enumerable.Range(1, 10)
                .Select(i => new
                {
                    Number = i,
                    NumberString = string.Concat("Number", i.ToString())
                });

            var table = ws.FirstCell()
                .InsertTable(data)
                .SetShowTotalsRow();

            table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Count;

            table.DataRange.Rows()
                .OrderByDescending(r => r.RowNumber())
                .ToList()
                .ForEach(r => r.WorksheetRow().Delete());

            Assert.That(table.DataRange, Is.Null);
            Assert.Throws<EmptyTableException>(() => wb.SaveAs(ms));
        }

        [Test]
        public void Save_totals_row_label_cell_with_sst_id_matching_the_label()
        {
            // Issue #2602 test. The totals row  wasn't saved with compact SST ID from file, but with a memory SST that has holes.
            TestHelper.CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();
                ws.Cell("A1").Value = "Dummy1"; // First inserted text - index=0, reference count = 1
                ws.Cell("A2").Value = "Dummy2"; // Second inserted text - index=1, reference count = 1
                ws.Cell("A3").Value = "Dummy3"; // Third inserted text - index=2, reference count = 1
                ws.Cell("A4").Value = "Text"; // Fourth inserted text - index=3, reference count = 1
                var table = ws.Cell("A5").InsertTable(new [] { ("Text", 17) }); // Also inserts header Item1 and Item2.
                table.ShowTotalsRow = true;
                table.Field(0).TotalsRowLabel = "Text"; // reference count = 3

                // Remove "Dummy*" text. That way, the "Text", "Item1" and "Item2" will be in index 0..2 that were occupied by Dummy*
                // Ensure that cell in total row label A7 references "Text" SST ID
                ws.Range("A1:A3").Value = Blank.Value;
            }, @"Other\Tables\TotalRowSstId.xlsx");
        }

        [Test]
        public void CanCreateTableWithWhiteSpaceColumnHeaders()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue("Header");
            ws.Cell("B1").SetValue(new string(' ', 1));
            ws.Cell("C1").SetValue(new string(' ', 2));
            ws.Cell("D1").SetValue(new string(' ', 3));

            var table = ws.Range("A1:E3").CreateTable("Table1");

            Assert.Multiple(() =>
            {
                Assert.That(table.Field(0).Name, Is.EqualTo("Header"));
                Assert.That(table.Field(1).Name, Is.EqualTo(new string(' ', 1)));
                Assert.That(table.Field(2).Name, Is.EqualTo(new string(' ', 2)));
                Assert.That(table.Field(3).Name, Is.EqualTo(new string(' ', 3)));
                Assert.That(table.Field(4).Name, Is.EqualTo("Column5"));
            });
        }

        [Test]
        public void TableNotFound()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Table("dummy"));
            Assert.Throws<ArgumentOutOfRangeException>(() => wb.Table("dummy"));
        }

        [Test]
        public void SecondTableOnNewSheetHasUniqueName()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var t1 = ws1.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
            Assert.That(t1.Name, Is.EqualTo("Table1"));

            var ws2 = wb.AddWorksheet();
            var t2 = ws2.FirstCell().InsertTable(Enumerable.Range(1, 10).Select(i => new { Number = i }));
            Assert.That(t2.Name, Is.EqualTo("Table2"));
        }

        private void AssertTablesAreEqual(IXLTable table1, IXLTable table2)
        {
            Assert.Multiple(() =>
            {
                Assert.That(table2.RangeAddress.ToString(XLReferenceStyle.A1, false), Is.EqualTo(table1.RangeAddress.ToString(XLReferenceStyle.A1, false)));
                Assert.That(table2.Fields.Count(), Is.EqualTo(table1.Fields.Count()));
            });
            for (int j = 0; j < table1.Fields.Count(); j++)
            {
                var originalField = table1.Fields.ElementAt(j);
                var copyField = table2.Fields.ElementAt(j);
                Assert.That(copyField.Name, Is.EqualTo(originalField.Name));
                if (table1.ShowTotalsRow)
                {
                    Assert.Multiple(() =>
                    {
                        Assert.That(copyField.TotalsRowFormulaA1, Is.EqualTo(originalField.TotalsRowFormulaA1));
                        Assert.That(copyField.TotalsRowFunction, Is.EqualTo(originalField.TotalsRowFunction));
                    });
                }
            }

            Assert.Multiple(() =>
            {
                Assert.That(table2.Name, Is.EqualTo(table1.Name));
                Assert.That(table2.ShowAutoFilter, Is.EqualTo(table1.ShowAutoFilter));
                Assert.That(table2.ShowColumnStripes, Is.EqualTo(table1.ShowColumnStripes));
                Assert.That(table2.ShowHeaderRow, Is.EqualTo(table1.ShowHeaderRow));
                Assert.That(table2.ShowRowStripes, Is.EqualTo(table1.ShowRowStripes));
                Assert.That(table2.ShowTotalsRow, Is.EqualTo(table1.ShowTotalsRow));
                Assert.That((table2.Style as XLStyle).Value, Is.EqualTo((table1.Style as XLStyle).Value));
                Assert.That(table2.Theme, Is.EqualTo(table1.Theme));
            });
        }

        //TODO: Delete table (not underlying range)
    }
}
