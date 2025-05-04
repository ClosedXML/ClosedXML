using ClosedXML.Attributes;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Tables
{
    [TestFixture]
    public class AppendingAndReplacingTableDataTests
    {
        public class TestObjectWithoutAttributes
        {
            public string Column1 { get; set; }
            public string Column2 { get; set; }
        }

        public class Person
        {
            public int Age { get; set; }

            [XLColumn(Header = "Last name", Order = 2)]
            public string LastName { get; set; }

            [XLColumn(Header = "First name", Order = 1)]
            public string FirstName { get; set; }

            [XLColumn(Header = "Full name", Order = 0)]
            public string FullName { get => string.Concat(FirstName, " ", LastName); }

            [XLColumn(Order = 3)]
            public DateTime DateOfBirth { get; set; }

            [XLColumn(Header = "Is active", Order = 4)]
            public bool IsActive;
        }

        private XLWorkbook PrepareWorkbook()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Tables");

            var data = new[]
            {
                new Person{FirstName = "Francois", LastName = "Botha", Age = 39, DateOfBirth = new DateTime(1980,1,1), IsActive = true},
                new Person{FirstName = "Leon", LastName = "Oosthuizen", Age = 40, DateOfBirth = new DateTime(1979,1,1), IsActive = false},
                new Person{FirstName = "Rian", LastName = "Prinsloo", Age = 41, DateOfBirth = new DateTime(1978,1,1), IsActive = false}
            };

            ws.FirstCell().CellRight().CellBelow().InsertTable(data);

            ws.Columns().AdjustToContents();

            return wb;
        }

        private XLWorkbook PrepareWorkbookWithAdditionalColumns()
        {
            var wb = PrepareWorkbook();
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();
            table.HeadersRow()
                .LastCell().CellRight()
                .InsertData(new[] { "CumulativeAge", "NameLength", "IsOld", "HardCodedValue" }, transpose: true);

            table.Resize(ws.Range(table.FirstCell(), table.LastCell().CellRight(4)));

            table.Field("CumulativeAge").DataCells.ForEach(c => c.FormulaA1 = $"SUM($G$3:G{c.WorksheetRow().RowNumber()})");
            table.Field("NameLength").DataCells.ForEach(c => c.FormulaA1 = $"LEN(B{c.WorksheetRow().RowNumber()})");
            table.Field("IsOld").DataCells.ForEach(c => c.FormulaA1 = $"=G{c.WorksheetRow().RowNumber()}>=40");
            table.Field("HardCodedValue").DataCells.Value = "40 is not old!";

            return wb;
        }

        private Person[] NewData
        {
            get
            {
                return new[]
                {
                    new Person{FirstName = "Michelle", LastName = "de Beer", Age = 35, DateOfBirth = new DateTime(1983,1,1), IsActive = false},
                    new Person{FirstName = "Marichen", LastName = "van der Gryp", Age = 30, DateOfBirth = new DateTime(1990,1,1), IsActive = true}
                };
            }
        }

        [Test]
        public void AddingEmptyEnumerables()
        {
            using var wb = PrepareWorkbook();
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = null;
            Assert.That(table.AppendData(personEnumerable), Is.Null);

            personEnumerable = new Person[] { };
            Assert.That(table.AppendData(personEnumerable), Is.Null);

            IEnumerable enumerable = null;
            Assert.That(table.AppendData(enumerable), Is.Null);

            enumerable = new Person[] { };
            Assert.That(table.AppendData(enumerable), Is.Null);
        }

        [Test]
        public void ReplaceWithEmptyEnumerables()
        {
            using var wb = PrepareWorkbook();
            var ws = wb.Worksheets.First();

            var table = ws.Tables.First();

            IEnumerable<Person> personEnumerable = null;
            Assert.Throws<InvalidOperationException>(() => table.ReplaceData(personEnumerable));

            personEnumerable = new Person[] { };
            Assert.Throws<InvalidOperationException>(() => table.ReplaceData(personEnumerable));

            IEnumerable enumerable = null;
            Assert.Throws<InvalidOperationException>(() => table.ReplaceData(enumerable));

            enumerable = new Person[] { };
            Assert.Throws<InvalidOperationException>(() => table.ReplaceData(enumerable));
        }

        [Test]
        public void CanAppendTypedEnumerable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                var addedRange = table.AppendData(personEnumerable);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G7"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(5));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanAppendToTableWithTotalsRow()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                IEnumerable<Person> personEnumerable = NewData;
                var addedRange = table.AppendData(personEnumerable);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G7"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(5));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanAppendTypedEnumerableAndPushDownCellsBelowTable()
        {
            using var ms = new MemoryStream();
            var value = "Some value that will be overwritten";
            IXLAddress address;
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                var cell = table.LastRow().FirstCell().CellRight(2).CellBelow(1);
                address = cell.Address;
                cell.Value = value;

                IEnumerable<Person> personEnumerable = NewData;
                var addedRange = table.AppendData(personEnumerable);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G7"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                var cell = ws.Cell(address);
                Assert.Multiple(() =>
                {
                    Assert.That(cell.Value, Is.EqualTo("de Beer"));
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(5));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));

                    Assert.That(cell.CellBelow(NewData.Length).Value, Is.EqualTo(value));
                });
            }
        }

        [Test]
        public void CanAppendUntypedEnumerable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                var list = new ArrayList();
                list.AddRange(NewData);

                var addedRange = table.AppendData(list);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G7"));

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(5));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanAppendDataTable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;

                var ws2 = wb.AddWorksheet("temp");
                var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

                var addedRange = table.AppendData(dataTable);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G7"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(5));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceWithTypedEnumerable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                var replacedRange = table.ReplaceData(personEnumerable);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G4"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(2));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceWithUntypedEnumerable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                var list = new ArrayList();
                list.AddRange(NewData);

                var replacedRange = table.ReplaceData(list);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G4"));

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(2));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceWithDataTable()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;

                var ws2 = wb.AddWorksheet("temp");
                var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

                var replacedRange = table.ReplaceData(dataTable);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G4"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(2));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceToTableWithTablesRow1()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                // Will cause table to overflow
                IEnumerable<Person> personEnumerable = NewData.Union(NewData).Union(NewData);
                var replacedRange = table.ReplaceData(personEnumerable);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G8"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(6));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceToTableWithTablesRow2()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();
                table.SetShowTotalsRow(true);
                table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                // Will cause table to shrink
                IEnumerable<Person> personEnumerable = NewData.Take(1);
                var replacedRange = table.ReplaceData(personEnumerable);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G3"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(1));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanReplaceWithUntypedEnumerableAndPropagateExtraColumns()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbookWithAdditionalColumns())
            {
                var ws = wb.Worksheets.First();
                var table = ws.Tables.First();

                var list = new ArrayList();
                list.AddRange(NewData);
                list.AddRange(NewData);

                var replacedRange = table.ReplaceData(list, propagateExtraColumns: true);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G6"));

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(4));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(10));

                    Assert.That(table.Worksheet.Cell("H5").FormulaA1, Is.EqualTo("SUM($G$3:G5)"));
                    Assert.That(table.Worksheet.Cell("H6").FormulaA1, Is.EqualTo("SUM($G$3:G6)"));
                    Assert.That(table.Worksheet.Cell("H5").Value, Is.EqualTo(100));
                    Assert.That(table.Worksheet.Cell("H6").Value, Is.EqualTo(130));

                    Assert.That(table.Worksheet.Cell("I5").FormulaA1, Is.EqualTo("LEN(B5)"));
                    Assert.That(table.Worksheet.Cell("I6").FormulaA1, Is.EqualTo("LEN(B6)"));
                    Assert.That(table.Worksheet.Cell("I5").Value, Is.EqualTo(16));
                    Assert.That(table.Worksheet.Cell("I6").Value, Is.EqualTo(21));

                    Assert.That(table.Worksheet.Cell("J5").FormulaA1, Is.EqualTo("G5>=40"));
                    Assert.That(table.Worksheet.Cell("J6").FormulaA1, Is.EqualTo("G6>=40"));
                    Assert.That(table.Worksheet.Cell("J5").Value, Is.EqualTo(false));
                    Assert.That(table.Worksheet.Cell("J6").Value, Is.EqualTo(false));

                    Assert.That(table.Worksheet.Cell("K5").Value, Is.EqualTo("40 is not old!"));
                    Assert.That(table.Worksheet.Cell("K6").Value, Is.EqualTo("40 is not old!"));
                });
            }
        }

        [Test]
        public void CanReplaceWithTypedEnumerableAndPropagateExtraColumns()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbookWithAdditionalColumns())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData.Concat(NewData).OrderBy(p => p.Age);
                var replacedRange = table.ReplaceData(personEnumerable, propagateExtraColumns: true);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G6"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(4));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(10));

                    Assert.That(table.Worksheet.Cell("H5").FormulaA1, Is.EqualTo("SUM($G$3:G5)"));
                    Assert.That(table.Worksheet.Cell("H6").FormulaA1, Is.EqualTo("SUM($G$3:G6)"));
                    Assert.That(table.Worksheet.Cell("H5").Value, Is.EqualTo(95));
                    Assert.That(table.Worksheet.Cell("H6").Value, Is.EqualTo(130));

                    Assert.That(table.Worksheet.Cell("I5").FormulaA1, Is.EqualTo("LEN(B5)"));
                    Assert.That(table.Worksheet.Cell("I6").FormulaA1, Is.EqualTo("LEN(B6)"));
                    Assert.That(table.Worksheet.Cell("I5").Value, Is.EqualTo(16));
                    Assert.That(table.Worksheet.Cell("I6").Value, Is.EqualTo(16));

                    Assert.That(table.Worksheet.Cell("J5").FormulaA1, Is.EqualTo("G5>=40"));
                    Assert.That(table.Worksheet.Cell("J6").FormulaA1, Is.EqualTo("G6>=40"));
                    Assert.That(table.Worksheet.Cell("J5").Value, Is.EqualTo(false));
                    Assert.That(table.Worksheet.Cell("J6").Value, Is.EqualTo(false));

                    Assert.That(table.Worksheet.Cell("K5").Value, Is.EqualTo("40 is not old!"));
                    Assert.That(table.Worksheet.Cell("K6").Value, Is.EqualTo("40 is not old!"));
                });
            }
        }

        [TestCase("ListOfPeople[Age]")] // Defined name formula without a A1 reference
        [TestCase("ListOfPeople!A1")] // Defined name formula with an A1 reference
        public void CanReplaceTableDataWhenWorksheetHasDefinedNames(string nameFormula)
        {
            // When table data are replaced, the size of a table is modified. That
            // means rows below it are shifted up/down and defined names should be
            // adjusted.
            // TODO: add assert for name shift when formulas are properly shifted. Originally, it threw even on defined name with A1 reference
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                ws.DefinedNames.Add("ListOfPeople_Age", nameFormula);

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = NewData;
                var replacedRange = table.ReplaceData(personEnumerable);

                Assert.That(replacedRange.RangeAddress.ToString(), Is.EqualTo("B3:G4"));

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(2));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(6));
                });
            }
        }

        [Test]
        public void CanAppendWithUntypedEnumerableAndPropagateExtraColumns()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbookWithAdditionalColumns())
            {
                var ws = wb.Worksheets.First();
                var table = ws.Tables.First();

                var list = new ArrayList();
                list.AddRange(NewData);
                list.AddRange(NewData);

                var appendedRange = table.AppendData(list, propagateExtraColumns: true);

                Assert.That(appendedRange.RangeAddress.ToString(), Is.EqualTo("B6:G9"));

                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(7));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(10));

                    Assert.That(table.Worksheet.Cell("H8").FormulaA1, Is.EqualTo("SUM($G$3:G8)"));
                    Assert.That(table.Worksheet.Cell("H9").FormulaA1, Is.EqualTo("SUM($G$3:G9)"));
                    Assert.That(table.Worksheet.Cell("H8").Value, Is.EqualTo(220));
                    Assert.That(table.Worksheet.Cell("H9").Value, Is.EqualTo(250));

                    Assert.That(table.Worksheet.Cell("I8").FormulaA1, Is.EqualTo("LEN(B8)"));
                    Assert.That(table.Worksheet.Cell("I9").FormulaA1, Is.EqualTo("LEN(B9)"));
                    Assert.That(table.Worksheet.Cell("I8").Value, Is.EqualTo(16));
                    Assert.That(table.Worksheet.Cell("I9").Value, Is.EqualTo(21));

                    Assert.That(table.Worksheet.Cell("J8").FormulaA1, Is.EqualTo("G8>=40"));
                    Assert.That(table.Worksheet.Cell("J9").FormulaA1, Is.EqualTo("G9>=40"));
                    Assert.That(table.Worksheet.Cell("J8").Value, Is.EqualTo(false));
                    Assert.That(table.Worksheet.Cell("J9").Value, Is.EqualTo(false));

                    Assert.That(table.Worksheet.Cell("K8").Value, Is.EqualTo("40 is not old!"));
                    Assert.That(table.Worksheet.Cell("K9").Value, Is.EqualTo("40 is not old!"));
                });
            }
        }

        [Test]
        public void CanAppendTypedEnumerableAndPropagateExtraColumns()
        {
            using var ms = new MemoryStream();
            using (var wb = PrepareWorkbookWithAdditionalColumns())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable =
                    NewData
                        .Concat(NewData)
                        .Concat(NewData)
                        .OrderBy(p => p.FirstName);

                var addedRange = table.AppendData(personEnumerable);

                Assert.That(addedRange.RangeAddress.ToString(), Is.EqualTo("B6:G11"));
                ws.Columns().AdjustToContents();

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                Assert.Multiple(() =>
                {
                    Assert.That(table.DataRange.RowCount(), Is.EqualTo(9));
                    Assert.That(table.DataRange.ColumnCount(), Is.EqualTo(10));

                    Assert.That(table.Worksheet.Cell("H10").FormulaA1, Is.EqualTo("SUM($G$3:G10)"));
                    Assert.That(table.Worksheet.Cell("H11").FormulaA1, Is.EqualTo("SUM($G$3:G11)"));
                    Assert.That(table.Worksheet.Cell("H10").Value, Is.EqualTo(280));
                    Assert.That(table.Worksheet.Cell("H11").Value, Is.EqualTo(315));

                    Assert.That(table.Worksheet.Cell("I10").FormulaA1, Is.EqualTo("LEN(B10)"));
                    Assert.That(table.Worksheet.Cell("I11").FormulaA1, Is.EqualTo("LEN(B11)"));
                    Assert.That(table.Worksheet.Cell("I10").Value, Is.EqualTo(16));
                    Assert.That(table.Worksheet.Cell("I11").Value, Is.EqualTo(16));

                    Assert.That(table.Worksheet.Cell("J10").FormulaA1, Is.EqualTo("G10>=40"));
                    Assert.That(table.Worksheet.Cell("J11").FormulaA1, Is.EqualTo("G11>=40"));
                    Assert.That(table.Worksheet.Cell("J10").Value, Is.EqualTo(false));
                    Assert.That(table.Worksheet.Cell("J11").Value, Is.EqualTo(false));

                    Assert.That(table.Worksheet.Cell("K10").Value, Is.EqualTo("40 is not old!"));
                    Assert.That(table.Worksheet.Cell("K11").Value, Is.EqualTo("40 is not old!"));
                });
            }
        }
    }
}
