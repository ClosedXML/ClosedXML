using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class AutoFilterTests
    {
        [Test]
        public void AutoFilterExpandsWithTable()
        {
            using (var wb = new XLWorkbook())
            {
                using (IXLWorksheet ws = wb.Worksheets.Add("Sheet1"))
                {
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
        }

        [Test]
        public void AutoFilterSortWhenNotInFirstRow()
        {
            using (var wb = new XLWorkbook())
            {
                using (IXLWorksheet ws = wb.Worksheets.Add("Sheet1"))
                {
                    ws.Cell(3, 3).SetValue("Names")
                        .CellBelow().SetValue("Manuel")
                        .CellBelow().SetValue("Carlos")
                        .CellBelow().SetValue("Dominic");
                    ws.RangeUsed().SetAutoFilter().Sort();
                    Assert.AreEqual("Carlos", ws.Cell(4, 3).GetString());
                }
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
            Assert.That(!ws.AutoFilter.Enabled);

            ws.RangeUsed().SetAutoFilter();
            Assert.That(ws.AutoFilter.Enabled);

            ws.AutoFilter.Clear();
            Assert.That(!ws.AutoFilter.Enabled);
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
                Assert.That(!ws.AutoFilter.Enabled);

                ws.RangeUsed().SetAutoFilter();
                Assert.That(ws.AutoFilter.Enabled);

                ws.RangeUsed().SetAutoFilter(false);
                Assert.That(!ws.AutoFilter.Enabled);
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
                    Assert.IsTrue(wb2.Worksheets.First().AutoFilter.Enabled);
                }
            }
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
                using (var ws = wb.Worksheets.Add("Sheet1"))
                {
                    ws.Cell(3, 3).SetValue("Names")
                        .CellBelow().SetValue("Manuel")
                        .CellBelow().SetValue("Carlos")
                        .CellBelow().SetValue("Dominic");

                    var autoFilter = ws.RangeUsed()
                        .SetAutoFilter();

                    autoFilter.Column(1).AddFilter("Carlos");

                    Assert.AreEqual("Carlos", ws.Cell(5, 3).GetString());
                    Assert.AreEqual(2, autoFilter.VisibleRows.Count());
                    Assert.AreEqual(3, autoFilter.VisibleRows.First().WorksheetRow().RowNumber());
                    Assert.AreEqual(5, autoFilter.VisibleRows.Last().WorksheetRow().RowNumber());
                }
            }
        }

        [Test]
        public void ReapplyAutoFilter()
        {
            using (var wb = new XLWorkbook())
            {
                using (var ws = wb.Worksheets.Add("Sheet1"))
                {
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
        }
    }
}
