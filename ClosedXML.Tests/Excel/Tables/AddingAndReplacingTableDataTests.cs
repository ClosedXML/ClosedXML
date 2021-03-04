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
            public String Column1 { get; set; }
            public String Column2 { get; set; }
        }

        public class Person
        {
            public int Age { get; set; }

            [XLColumn(Header = "Last name", Order = 2)]
            public String LastName { get; set; }

            [XLColumn(Header = "First name", Order = 1)]
            public String FirstName { get; set; }

            [XLColumn(Header = "Full name", Order = 0)]
            public String FullName { get => string.Concat(FirstName, " ", LastName); }

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
            using (var wb = PrepareWorkbook())
            {
                var ws = wb.Worksheets.First();

                var table = ws.Tables.First();

                IEnumerable<Person> personEnumerable = null;
                Assert.AreEqual(null, table.AppendData(personEnumerable));

                personEnumerable = new Person[] { };
                Assert.AreEqual(null, table.AppendData(personEnumerable));

                IEnumerable enumerable = null;
                Assert.AreEqual(null, table.AppendData(enumerable));

                enumerable = new Person[] { };
                Assert.AreEqual(null, table.AppendData(enumerable));
            }
        }

        [Test]
        public void ReplaceWithEmptyEnumerables()
        {
            using (var wb = PrepareWorkbook())
            {
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
        }

        [Test]
        public void CanAppendTypedEnumerable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    IEnumerable<Person> personEnumerable = NewData;
                    var addedRange = table.AppendData(personEnumerable);

                    Assert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(5, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanAppendToTableWithTotalsRow()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();
                    table.SetShowTotalsRow(true);
                    table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                    IEnumerable<Person> personEnumerable = NewData;
                    var addedRange = table.AppendData(personEnumerable);

                    Assert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(5, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanAppendTypedEnumerableAndPushDownCellsBelowTable()
        {
            using (var ms = new MemoryStream())
            {
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

                    Assert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    var cell = ws.Cell(address);
                    Assert.AreEqual("de Beer", cell.Value);
                    Assert.AreEqual(5, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());

                    Assert.AreEqual(value, cell.CellBelow(NewData.Count()).Value);
                }
            }
        }

        [Test]
        public void CanAppendUntypedEnumerable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    var list = new ArrayList();
                    list.AddRange(NewData);

                    var addedRange = table.AppendData(list);

                    Assert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());

                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(5, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanAppendDataTable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    IEnumerable<Person> personEnumerable = NewData;

                    var ws2 = wb.AddWorksheet("temp");
                    var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

                    var addedRange = table.AppendData(dataTable);

                    Assert.AreEqual("B6:G7", addedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(5, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceWithTypedEnumerable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    IEnumerable<Person> personEnumerable = NewData;
                    var replacedRange = table.ReplaceData(personEnumerable);

                    Assert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(2, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceWithUntypedEnumerable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    var list = new ArrayList();
                    list.AddRange(NewData);

                    var replacedRange = table.ReplaceData(list);

                    Assert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());

                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(2, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceWithDataTable()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    IEnumerable<Person> personEnumerable = NewData;

                    var ws2 = wb.AddWorksheet("temp");
                    var dataTable = ws2.FirstCell().InsertTable(personEnumerable).AsNativeDataTable();

                    var replacedRange = table.ReplaceData(dataTable);

                    Assert.AreEqual("B3:G4", replacedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(2, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceToTableWithTablesRow1()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();
                    table.SetShowTotalsRow(true);
                    table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                    // Will cause table to overflow
                    IEnumerable<Person> personEnumerable = NewData.Union(NewData).Union(NewData);
                    var replacedRange = table.ReplaceData(personEnumerable);

                    Assert.AreEqual("B3:G8", replacedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(6, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceToTableWithTablesRow2()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbook())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();
                    table.SetShowTotalsRow(true);
                    table.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Average;

                    // Will cause table to shrink
                    IEnumerable<Person> personEnumerable = NewData.Take(1);
                    var replacedRange = table.ReplaceData(personEnumerable);

                    Assert.AreEqual("B3:G3", replacedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(1, table.DataRange.RowCount());
                    Assert.AreEqual(6, table.DataRange.ColumnCount());
                }
            }
        }

        [Test]
        public void CanReplaceWithUntypedEnumerableAndPropagateExtraColumns()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbookWithAdditionalColumns())
                {
                    var ws = wb.Worksheets.First();
                    var table = ws.Tables.First();

                    var list = new ArrayList();
                    list.AddRange(NewData);
                    list.AddRange(NewData);

                    var replacedRange = table.ReplaceData(list, propagateExtraColumns: true);

                    Assert.AreEqual("B3:G6", replacedRange.RangeAddress.ToString());

                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(4, table.DataRange.RowCount());
                    Assert.AreEqual(10, table.DataRange.ColumnCount());

                    Assert.AreEqual("SUM($G$3:G5)", table.Worksheet.Cell("H5").FormulaA1);
                    Assert.AreEqual("SUM($G$3:G6)", table.Worksheet.Cell("H6").FormulaA1);
                    Assert.AreEqual(100, table.Worksheet.Cell("H5").Value);
                    Assert.AreEqual(130, table.Worksheet.Cell("H6").Value);

                    Assert.AreEqual("LEN(B5)", table.Worksheet.Cell("I5").FormulaA1);
                    Assert.AreEqual("LEN(B6)", table.Worksheet.Cell("I6").FormulaA1);
                    Assert.AreEqual(16, table.Worksheet.Cell("I5").Value);
                    Assert.AreEqual(21, table.Worksheet.Cell("I6").Value);

                    Assert.AreEqual("G5>=40", table.Worksheet.Cell("J5").FormulaA1);
                    Assert.AreEqual("G6>=40", table.Worksheet.Cell("J6").FormulaA1);
                    Assert.AreEqual(false, table.Worksheet.Cell("J5").Value);
                    Assert.AreEqual(false, table.Worksheet.Cell("J6").Value);

                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K5").Value);
                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K6").Value);
                }
            }
        }

        [Test]
        public void CanReplaceWithTypedEnumerableAndPropagateExtraColumns()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbookWithAdditionalColumns())
                {
                    var ws = wb.Worksheets.First();

                    var table = ws.Tables.First();

                    IEnumerable<Person> personEnumerable = NewData.Concat(NewData).OrderBy(p => p.Age);
                    var replacedRange = table.ReplaceData(personEnumerable, propagateExtraColumns: true);

                    Assert.AreEqual("B3:G6", replacedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(4, table.DataRange.RowCount());
                    Assert.AreEqual(10, table.DataRange.ColumnCount());

                    Assert.AreEqual("SUM($G$3:G5)", table.Worksheet.Cell("H5").FormulaA1);
                    Assert.AreEqual("SUM($G$3:G6)", table.Worksheet.Cell("H6").FormulaA1);
                    Assert.AreEqual(95, table.Worksheet.Cell("H5").Value);
                    Assert.AreEqual(130, table.Worksheet.Cell("H6").Value);

                    Assert.AreEqual("LEN(B5)", table.Worksheet.Cell("I5").FormulaA1);
                    Assert.AreEqual("LEN(B6)", table.Worksheet.Cell("I6").FormulaA1);
                    Assert.AreEqual(16, table.Worksheet.Cell("I5").Value);
                    Assert.AreEqual(16, table.Worksheet.Cell("I6").Value);

                    Assert.AreEqual("G5>=40", table.Worksheet.Cell("J5").FormulaA1);
                    Assert.AreEqual("G6>=40", table.Worksheet.Cell("J6").FormulaA1);
                    Assert.AreEqual(false, table.Worksheet.Cell("J5").Value);
                    Assert.AreEqual(false, table.Worksheet.Cell("J6").Value);

                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K5").Value);
                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K6").Value);
                }
            }
        }

        [Test]
        public void CanAppendWithUntypedEnumerableAndPropagateExtraColumns()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = PrepareWorkbookWithAdditionalColumns())
                {
                    var ws = wb.Worksheets.First();
                    var table = ws.Tables.First();

                    var list = new ArrayList();
                    list.AddRange(NewData);
                    list.AddRange(NewData);

                    var appendedRange = table.AppendData(list, propagateExtraColumns: true);

                    Assert.AreEqual("B6:G9", appendedRange.RangeAddress.ToString());

                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(7, table.DataRange.RowCount());
                    Assert.AreEqual(10, table.DataRange.ColumnCount());

                    Assert.AreEqual("SUM($G$3:G8)", table.Worksheet.Cell("H8").FormulaA1);
                    Assert.AreEqual("SUM($G$3:G9)", table.Worksheet.Cell("H9").FormulaA1);
                    Assert.AreEqual(220, table.Worksheet.Cell("H8").Value);
                    Assert.AreEqual(250, table.Worksheet.Cell("H9").Value);

                    Assert.AreEqual("LEN(B8)", table.Worksheet.Cell("I8").FormulaA1);
                    Assert.AreEqual("LEN(B9)", table.Worksheet.Cell("I9").FormulaA1);
                    Assert.AreEqual(16, table.Worksheet.Cell("I8").Value);
                    Assert.AreEqual(21, table.Worksheet.Cell("I9").Value);

                    Assert.AreEqual("G8>=40", table.Worksheet.Cell("J8").FormulaA1);
                    Assert.AreEqual("G9>=40", table.Worksheet.Cell("J9").FormulaA1);
                    Assert.AreEqual(false, table.Worksheet.Cell("J8").Value);
                    Assert.AreEqual(false, table.Worksheet.Cell("J9").Value);

                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K8").Value);
                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K9").Value);
                }
            }
        }

        [Test]
        public void CanAppendTypedEnumerableAndPropagateExtraColumns()
        {
            using (var ms = new MemoryStream())
            {
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

                    Assert.AreEqual("B6:G11", addedRange.RangeAddress.ToString());
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var table = wb.Worksheets.SelectMany(ws => ws.Tables).First();

                    Assert.AreEqual(9, table.DataRange.RowCount());
                    Assert.AreEqual(10, table.DataRange.ColumnCount());

                    Assert.AreEqual("SUM($G$3:G10)", table.Worksheet.Cell("H10").FormulaA1);
                    Assert.AreEqual("SUM($G$3:G11)", table.Worksheet.Cell("H11").FormulaA1);
                    Assert.AreEqual(280, table.Worksheet.Cell("H10").Value);
                    Assert.AreEqual(315, table.Worksheet.Cell("H11").Value);

                    Assert.AreEqual("LEN(B10)", table.Worksheet.Cell("I10").FormulaA1);
                    Assert.AreEqual("LEN(B11)", table.Worksheet.Cell("I11").FormulaA1);
                    Assert.AreEqual(16, table.Worksheet.Cell("I10").Value);
                    Assert.AreEqual(16, table.Worksheet.Cell("I11").Value);

                    Assert.AreEqual("G10>=40", table.Worksheet.Cell("J10").FormulaA1);
                    Assert.AreEqual("G11>=40", table.Worksheet.Cell("J11").FormulaA1);
                    Assert.AreEqual(false, table.Worksheet.Cell("J10").Value);
                    Assert.AreEqual(false, table.Worksheet.Cell("J11").Value);

                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K10").Value);
                    Assert.AreEqual("40 is not old!", table.Worksheet.Cell("K11").Value);
                }
            }
        }
    }
}
