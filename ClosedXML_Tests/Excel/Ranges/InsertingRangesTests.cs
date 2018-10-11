using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class InsertingRangesTests
    {
        [Test]
        public void InsertingColumnsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            IXLColumn column1 = ws.Column(1);
            column1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            IXLColumn column2 = ws.Column(2);
            column2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
            column2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

            column1.InsertColumnsAfter(1);
            column1.InsertColumnsBefore(1);
            column2.InsertColumnsBefore(1);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Column(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Column(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Xanadu, ws.Column(5).Style.Fill.BackgroundColor);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Cell(2, 1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(2, 5).Style.Fill.BackgroundColor);
        }

        [Test]
        public void InsertingRowsAbove()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            ws.Cell("B3").SetValue("X")
                .CellBelow().SetValue("B");

            IXLRangeRow r = ws.Range("B4").InsertRowsAbove(1).First();
            r.Cell(1).SetValue("A");

            Assert.AreEqual("X", ws.Cell("B3").GetString());
            Assert.AreEqual("A", ws.Cell("B4").GetString());
            Assert.AreEqual("B", ws.Cell("B5").GetString());
        }

        [Test]
        public void InsertingRowsPreservesFormatting()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            IXLRow row1 = ws.Row(1);
            row1.Style.Fill.SetBackgroundColor(XLColor.FrenchLilac);
            row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.Fulvous);
            IXLRow row2 = ws.Row(2);
            row2.Style.Fill.SetBackgroundColor(XLColor.Xanadu);
            row2.Cell(2).Style.Fill.SetBackgroundColor(XLColor.MacaroniAndCheese);

            row1.InsertRowsBelow(1);
            row1.InsertRowsAbove(1);
            row2.InsertRowsAbove(1);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FrenchLilac, ws.Row(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Xanadu, ws.Row(5).Style.Fill.BackgroundColor);

            Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Cell(1, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(2, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(3, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.Fulvous, ws.Cell(4, 2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.MacaroniAndCheese, ws.Cell(5, 2).Style.Fill.BackgroundColor);
        }

        [Test]
        public void InsertingRowsPreservesComments()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").SetValue("Insert Below");
            ws.Cell("A2").SetValue("Already existing cell");
            ws.Cell("A3").SetValue("Cell with comment").Comment.AddText("Comment here");

            ws.Row(1).InsertRowsBelow(2);
            Assert.AreEqual("Comment here", ws.Cell("A5").Comment.Text);
        }

        [Test]
        public void InsertingColumnsPreservesComments()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").SetValue("Insert to the right");
            ws.Cell("B1").SetValue("Already existing cell");
            ws.Cell("C1").SetValue("Cell with comment").Comment.AddText("Comment here");

            ws.Column(1).InsertColumnsAfter(2);
            Assert.AreEqual("Comment here", ws.Cell("E1").Comment.Text);
        }

        [Test]
        [TestCase("C4:F7", "C4:F7", 2, "E4:H7")] // Coincide, shift right
        [TestCase("C4:F7", "C4:F7", -2, "C4:D7")] // Coincide, shift left
        [TestCase("D5:E6", "C4:F7", 2, "F5:G6")] // Inside, shift right
        [TestCase("D5:E6", "C4:F7", -2, "C5:C6")] // Inside, shift left
        [TestCase("B4:G7", "C4:F7", 2, "B4:I7")] // Includes, shift right
        [TestCase("B4:G7", "C4:F7", -2, "B4:E7")] // Includes, shift left
        [TestCase("B4:E7", "C4:F7", 2, "B4:G7")] // Intersects at left, shift right
        [TestCase("B4:E7", "C4:F7", -2, "B4:C7")] // Intersects at left, shift left
        [TestCase("D4:G7", "C4:F7", 2, "F4:I7")] // Intersects at right, shift right
        [TestCase("D4:G7", "C4:F7", -2, "C4:E7")] // Intersects at right, shift left
        [TestCase("A5:B6", "C4:F7", 2, "A5:B6")] // No intersection, at left, shift right
        [TestCase("A5:B6", "C4:F7", -1, "A5:B6")] // No intersection, at left, shift left
        [TestCase("H5:I6", "C4:F7", 2, "J5:K6")] // No intersection, at right, shift right
        [TestCase("H5:I6", "C4:F7", -2, "F5:G6")] // No intersection, at right, shift left
        [TestCase("C8:F11", "C4:F7", 2, "C8:F11")] // Different rows
        [TestCase("B1:B8", "A1:C4", 1, "B1:B8")]  // More rows, shift right
        [TestCase("B1:B8", "A1:C4", -1, "B1:B8")]  // More rows, shift left
        public void ShiftColumnsValid(string thisRangeAddress, string shiftedRangeAddress, int shiftedColumns, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                var thisRange = ws.Range(thisRangeAddress) as XLRange;
                var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

                thisRange.WorksheetRangeShiftedColumns(shiftedRange, shiftedColumns);

                Assert.IsTrue(thisRange.RangeAddress.IsValid);
                Assert.AreEqual(expectedRange, thisRange.RangeAddress.ToString());
            }
        }

        [Test]
        [TestCase("B1:B4", "A1:C4", -2)] // Shift left too much
        public void ShiftColumnsInvalid(string thisRangeAddress, string shiftedRangeAddress, int shiftedColumns)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                var thisRange = ws.Range(thisRangeAddress) as XLRange;
                var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

                thisRange.WorksheetRangeShiftedColumns(shiftedRange, shiftedColumns);

                Assert.IsFalse(thisRange.RangeAddress.IsValid);
            }
        }

        [Test]
        [TestCase("C4:F7", "C4:F7", 2, "C6:F9")]   // Coincide, shift down
        [TestCase("C4:F7", "C4:F7", -2, "C4:F5")]   // Coincide, shift up
        [TestCase("D5:E6", "C4:F7", 2, "D7:E8")]   // Inside, shift down
        [TestCase("D5:E6", "C4:F7", -2, "D4:E4")]   // Inside, shift up
        [TestCase("C3:F8", "C4:F7", 2, "C3:F10")]  // Includes, shift down
        [TestCase("C3:F8", "C4:F7", -2, "C3:F6")]   // Includes, shift up
        [TestCase("C3:F6", "C4:F7", 2, "C3:F8")]   // Intersects at top, shift down
        [TestCase("C2:F6", "C4:F7", -3, "C2:F3")]   // Intersects at top, shift up to the sheet boundary
        [TestCase("C3:F6", "C4:F7", -2, "C3:F4")]   // Intersects at top, shift up
        [TestCase("C5:F8", "C4:F7", 2, "C7:F10")]  // Intersects at bottom, shift down
        [TestCase("C5:F8", "C4:F7", -2, "C4:F6")]   // Intersects at bottom, shift up
        [TestCase("C1:F3", "C4:F7", 2, "C1:F3")]   // No intersection, at top, shift down
        [TestCase("C1:F3", "C4:F7", -2, "C1:F3")]   // No intersection, at top, shift up
        [TestCase("C8:F10", "C4:F7", 2, "C10:F12")] // No intersection, at bottom, shift down
        [TestCase("C8:F10", "C4:F7", -2, "C6:F8")]   // No intersection, at bottom, shift up
        [TestCase("G4:J7", "C4:F7", 2, "G4:J7")]   // Different columns
        [TestCase("A2:D2", "A1:C4", 1, "A2:D2")]   // More columns, shift down
        [TestCase("A2:D2", "A1:C4", -1, "A2:D2")]   // More columns, shift up
        public void ShiftRowsValid(string thisRangeAddress, string shiftedRangeAddress, int shiftedRows, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                var thisRange = ws.Range(thisRangeAddress) as XLRange;
                var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

                thisRange.WorksheetRangeShiftedRows(shiftedRange, shiftedRows);

                Assert.IsTrue(thisRange.RangeAddress.IsValid);
                Assert.AreEqual(expectedRange, thisRange.RangeAddress.ToString());
            }
        }

        [Test]
        [TestCase("A2:C2", "A1:C4", -2)] // Shift up too much
        public void ShiftRowsInvalid(string thisRangeAddress, string shiftedRangeAddress, int shiftedRows)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                var thisRange = ws.Range(thisRangeAddress) as XLRange;
                var shiftedRange = ws.Range(shiftedRangeAddress) as XLRange;

                thisRange.WorksheetRangeShiftedRows(shiftedRange, shiftedRows);

                Assert.IsFalse(thisRange.RangeAddress.IsValid);
            }
        }

        [Test]
        public void InsertZeroColumnsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsAfter(0));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsBefore(0));
        }

        [Test]
        public void InsertNegativeNumberOfColumnsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsAfter(-1));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsBefore(-1));
        }

        [Test]
        public void InsertTooLargeNumberOfColumnsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsAfter(16385));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertColumnsBefore(16385));
        }

        [Test]
        public void InsertZeroRowsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsAbove(0));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsBelow(0));
        }

        [Test]
        public void InsertNegativeNumberOfRowsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsAbove(-1));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsBelow(-1));
        }

        [Test]
        public void InsertTooLargeNumberOrRowsFails()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var range = ws.FirstCell().AsRange();
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsAbove(1048577));
            Assert.Throws(typeof(ArgumentOutOfRangeException), () => range.InsertRowsBelow(1048577));
        }

        [Test]
        public void MergedRangesConsistencyWhenInsertingRows()
        {
            // https://github.com/ClosedXML/ClosedXML/issues/1013
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                //create merged row
                ws.Cell("A1").Value = "Merged Row(1) of Range (A1:F1)";
                ws.Range("A1:F1").Row(1).Merge();

                var row = ws.FirstRow();

                // Add some lines and copy format & merging
                for (var r = 1; r <= 10; r++)
                {
                    row.InsertRowsBelow(1);         // insert a row below row 1, as a row 2
                    row.CopyTo(row.RowBelow());     // copy format and merging from row 1 to row 2

                    var duplicates = ws.MergedRanges
                        .GroupBy(s => s.ToString())
                        .Where(g => g.Count() > 1)
                        .Select(y => new { Element = y.Key, Counter = y.Count() })
                        .ToList();

                    Assert.AreEqual(0, duplicates.Count);
                }
            }
        }
    }
}
