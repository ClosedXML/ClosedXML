using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class AutoFilterTests
    {
        [Test]
        public void AutoFilterExpandsWithTable()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.FirstCell().SetValue("Categories")
                    .CellBelow().SetValue("1")
                    .CellBelow().SetValue("2");

                IXLTable table = ws.RangeUsed().CreateTable();

                var listOfArr = new List<Int32>();
                listOfArr.Add(3);
                listOfArr.Add(4);
                listOfArr.Add(5);
                listOfArr.Add(6);

                table.DataRange.InsertRowsBelow(listOfArr.Count - table.DataRange.RowCount());
                table.DataRange.FirstCell().InsertData(listOfArr);

                Assert.AreEqual("A1:A5", table.AutoFilter.Range.RangeAddress.ToStringRelative());
                Assert.AreEqual(5, table.AutoFilter.VisibleRows.Count());
            }
        }

        [Test]
        public void AutoFilterSortWhenNotInFirstRow()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell(3, 3).SetValue("Names")
                    .CellBelow().SetValue("Manuel")
                    .CellBelow().SetValue("Carlos")
                    .CellBelow().SetValue("Dominic");
                ws.RangeUsed().SetAutoFilter().Sort();
                Assert.AreEqual("Carlos", ws.Cell(4, 3).GetText());
            }
        }

        [Test]
        public void CanClearAutoFilter()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.AutoFilter.Clear(); // We should be able to clear a filter even if it hasn't been set.
            Assert.That(!ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter();
            Assert.That(ws.AutoFilter.IsEnabled);

            ws.AutoFilter.Clear();
            Assert.That(!ws.AutoFilter.IsEnabled);
        }

        [Test]
        public void CanClearAutoFilter2()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("AutoFilter");
                ws.Cell("A1").Value = "Names";
                ws.Cell("A2").Value = "John";
                ws.Cell("A3").Value = "Hank";
                ws.Cell("A4").Value = "Dagny";

                ws.SetAutoFilter(false);
                Assert.That(!ws.AutoFilter.IsEnabled);

                ws.RangeUsed().SetAutoFilter();
                Assert.That(ws.AutoFilter.IsEnabled);

                ws.RangeUsed().SetAutoFilter(false);
                Assert.That(!ws.AutoFilter.IsEnabled);
            }
        }

        [Test]
        public void CanCopyAutoFilterToNewSheetOnNewWorkbook()
        {
            using (var ms1 = new MemoryStream())
            using (var ms2 = new MemoryStream())
            {
                using (var wb1 = new XLWorkbook())
                using (var wb2 = new XLWorkbook())
                {
                    var ws = wb1.Worksheets.Add("AutoFilter");
                    ws.Cell("A1").Value = "Names";
                    ws.Cell("A2").Value = "John";
                    ws.Cell("A3").Value = "Hank";
                    ws.Cell("A4").Value = "Dagny";

                    ws.RangeUsed().SetAutoFilter();

                    wb1.SaveAs(ms1);

                    ws.CopyTo(wb2, ws.Name);
                    wb2.SaveAs(ms2);
                }

                using (var wb2 = new XLWorkbook(ms2))
                {
                    Assert.IsTrue(wb2.Worksheets.First().AutoFilter.IsEnabled);
                }
            }
        }

        [Test]
        public void CannotAddAutoFilterOverExistingTable()
        {
            using var wb = new XLWorkbook();

            var data = Enumerable.Range(1, 10).Select(i => new
            {
                Index = i,
                String = $"String {i}"
            });

            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertTable(data);

            Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().SetAutoFilter());
        }

        [Test]
        [TestCase("A1:A4")]
        [TestCase("A1:B4")]
        [TestCase("A1:C4")]
        public void AutoFilterRangeRemainsValidOnInsertColumn(string rangeAddress)
        {
            //Arrange
            using (var ms1 = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("AutoFilter");
                    ws.Cell("A1").Value = "Ids";
                    ws.Cell("B1").Value = "Names";
                    ws.Cell("B2").Value = "John";
                    ws.Cell("B3").Value = "Hank";
                    ws.Cell("B4").Value = "Dagny";
                    ws.Cell("C1").Value = "Phones";

                    ws.Range("B1:B4").SetAutoFilter(true);

                    //Act
                    var range = ws.Range(rangeAddress);
                    range.InsertColumnsBefore(1);

                    //Assert
                    Assert.IsTrue(ws.AutoFilter.Range.RangeAddress.IsValid);
                }
            }
        }

        [Test]
        public void AutoFilterVisibleRows()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell(3, 3).SetValue("Names")
                    .CellBelow().SetValue("Manuel")
                    .CellBelow().SetValue("Carlos")
                    .CellBelow().SetValue("Dominic");

                var autoFilter = ws.RangeUsed()
                    .SetAutoFilter();

                autoFilter.Column(1).AddFilter("Carlos");

                Assert.AreEqual("Carlos", ws.Cell(5, 3).GetText());
                Assert.AreEqual(2, autoFilter.VisibleRows.Count());
                Assert.AreEqual(3, autoFilter.VisibleRows.First().WorksheetRow().RowNumber());
                Assert.AreEqual(5, autoFilter.VisibleRows.Last().WorksheetRow().RowNumber());
            }
        }

        [Test]
        public void ReapplyAutoFilter()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell(3, 3).SetValue("Names")
                    .CellBelow().SetValue("Manuel")
                    .CellBelow().SetValue("Carlos")
                    .CellBelow().SetValue("Dominic")
                    .CellBelow().SetValue("Jose");

                var autoFilter = ws.RangeUsed()
                    .SetAutoFilter();

                autoFilter.Column(1).AddFilter("Carlos");

                Assert.AreEqual(3, autoFilter.HiddenRows.Count());

                // Unhide the rows so that the table is out of sync with the filter
                autoFilter.HiddenRows.ForEach(r => r.WorksheetRow().Unhide());
                Assert.False(autoFilter.HiddenRows.Any());

                autoFilter.Reapply();
                Assert.AreEqual(3, autoFilter.HiddenRows.Count());
            }
        }

        [Test]
        public void CanLoadAutoFilterWithThousandsSeparator()
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                // Set thread culture to French, which should format numbers using a space as thousands separator
                var culture = CultureInfo.CreateSpecificCulture("fr-FR");

                // The value in sheet that will be compared with autofilter value is a number
                // `10000`. That number will be formatted using culture to `10 000.00` thanks to
                // modified properties of culture - period instead of a comma for decimal separator
                // and space as group separator. The formatted number will thus match with the
                // filter value.
                culture.NumberFormat.NumberDecimalSeparator = ".";
                culture.NumberFormat.NumberGroupSeparator = " ";

                Thread.CurrentThread.CurrentCulture = culture;

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();

                    // Regular filter compares values as strings, doesn't convert to XLCellValue,
                    // so the value is read from the file as a text despite looking like a number.
                    Assert.AreEqual("10 000.00", ((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value);
                    Assert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());

                    ws.AutoFilter.Reapply();
                    Assert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());
                }

                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual("10 000.00", ((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value);

                    var v = ws.AutoFilter.VisibleRows.Select(r => r.FirstCell().Value).ToList();
                    Assert.AreEqual(2, ws.AutoFilter.VisibleRows.Count());

                    ws.AutoFilter.Reapply();
                    Assert.AreEqual(1, ws.AutoFilter.VisibleRows.Count());
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }
        }

        [Test]
        public void Issue1917NotContainsFilter()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Test");
                    ws.Cell(1, 1).SetValue("StringCol");

                    for (var i = 0; i < 5; i++)
                    {
                        ws.Cell(i + 2, 1).SetValue($"String{i}");
                    }

                    var autoFilter = ws.RangeUsed()
                        .SetAutoFilter();

                    autoFilter.Column(1).NotContains("String3");
                    Assert.AreEqual(1, autoFilter.HiddenRows.Count());

                    wb.SaveAs(ms);
                }

                ms.Position = 0;
                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.Worksheet("Test");
                    var autoFilter = ws.AutoFilter;

                    autoFilter.Reapply();
                    Assert.AreEqual(1, autoFilter.HiddenRows.Count());
                }
            }
        }

        [Test]
        [TestCase("ends")]
        [TestCase("begins")]
        [TestCase("equal")]
        [TestCase("contains")]
        public void NotStringFilter(string type)
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Test");
                    ws.Cell(1, 1).SetValue("StringCol");

                    for (var i = 0; i < 5; i++)
                    {
                        ws.Cell(i + 2, 1).SetValue($"{i}-String{i}");
                    }

                    ws.Columns().AdjustToContents();
                    var autoFilter = ws.RangeUsed()
                        .SetAutoFilter();

                    switch (type)
                    {
                        case "ends":
                            autoFilter.Column(1).NotEndsWith("3");
                            break;
                        case "begins":
                            autoFilter.Column(1).NotBeginsWith("3");
                            break;
                        case "equal":
                            autoFilter.Column(1).NotEqualTo("3-String3");
                            break;
                        case "contains":
                            autoFilter.Column(1).NotContains("3-");
                            break;
                    }
                    Assert.AreEqual(1, autoFilter.HiddenRows.Count());

                    wb.SaveAs(ms);
                }

                ms.Position = 0;
                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.Worksheet("Test");
                    var autoFilter = ws.AutoFilter;

                    autoFilter.Reapply();
                    Assert.AreEqual(1, autoFilter.HiddenRows.Count());
                }
            }
        }
    }
}
