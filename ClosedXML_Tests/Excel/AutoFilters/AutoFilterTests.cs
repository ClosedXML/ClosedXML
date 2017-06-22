using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

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
                    table.DataRange.FirstCell().InsertData(listOfArr.AsEnumerable());

                    Assert.AreEqual("A1:A5", table.AutoFilter.Range.RangeAddress.ToStringRelative());
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
                    Assert.AreEqual(ws.Cell(4, 3).GetString(), "Carlos");
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
    }
}