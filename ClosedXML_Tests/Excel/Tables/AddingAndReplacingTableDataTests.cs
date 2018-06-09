using ClosedXML.Attributes;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel.Tables
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
    }
}
