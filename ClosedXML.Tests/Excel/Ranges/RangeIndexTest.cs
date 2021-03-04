using ClosedXML.Excel;
using ClosedXML.Excel.Patterns;
using ClosedXML.Excel.Ranges.Index;
using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges
{
    [TestFixture]
    public class RangeIndexTest
    {
        private const int TEST_COUNT = 10000;

        [Test]
        public void FindExistingMatches()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    for (int j = 2; j <= 4; j++)
                    {
                        var address = new XLAddress(ws, i * 2, j, false, false);
                        Assert.True(index.Contains(in address));
                    }
                }
            }
        }

        [Test]
        public void FindNonExistingMatches()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    var address = new XLAddress(ws, i * 2 + 1, 3, false, false);
                    Assert.False(index.Contains(in address));
                }
            }
        }

        [Test]
        public void FindExistingIntersections()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    var rangeAddress = new XLRangeAddress(
                        new XLAddress(ws, i * 2, 1 + i % 4, false, false),
                        new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false));

                    Assert.True(index.Intersects(in rangeAddress));
                }

                for (int i = 2; i < 4; i++)
                {
                    var columnAddress = XLRangeAddress.EntireColumn(ws, i);
                    Assert.True(index.Intersects(in columnAddress));
                }
            }
        }

        [Test]
        public void FindNonExistingIntersections()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    var rangeAddress = new XLRangeAddress(
                        new XLAddress(ws, i * 2 + 1, 1 + i % 4, false, false),
                        new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false));

                    Assert.False(index.Intersects(in rangeAddress));
                }

                var columnAddress = XLRangeAddress.EntireColumn(ws, 1);
                Assert.False(index.Intersects(in columnAddress));
                columnAddress = XLRangeAddress.EntireColumn(ws, 5);
                Assert.False(index.Intersects(in columnAddress));
            }
        }

        [Test]
        public void FindMatchAfterColumnShifting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                ws.Column(1).InsertColumnsBefore(1000);

                var address = new XLAddress(ws, 102, 1004, false, false);

                Assert.True(index.Contains(in address));
            }
        }

        [Test]
        public void FindIntersectionsAfterColumnShifting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                ws.Column(3).InsertColumnsBefore(2);

                var rangeAddress = new XLRangeAddress(ws, "F102:E103");

                Assert.True(index.Intersects(in rangeAddress));
            }
        }

        [Test]
        public void FindMatchAfterRowShifting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                ws.Row(10).InsertRowsBelow(3);

                var address = new XLAddress(ws, 103, 4, false, false);

                Assert.True(index.Contains(in address));
            }
        }

        [Test]
        public void FindIntersectionsAfterRowShifting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var index = FillIndexWithTestData(ws);

                ws.Row(10).InsertRowsBelow(3);

                var rangeAddress = new XLRangeAddress(ws, "C103:E103");

                Assert.True(index.Intersects(in rangeAddress));
            }
        }

        [Test]
        public void CreateQuadTree()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var quadTree = new Quadrant();
                var range = ws.Range("BT76:CA87");

                quadTree.Add(range);

                var level0 = quadTree;
                Assert.AreEqual(1, level0.MinimumColumn);
                Assert.AreEqual(XLHelper.MaxColumnNumber, level0.MaximumColumn);
                Assert.AreEqual(1, level0.MinimumRow);
                Assert.AreEqual(XLHelper.MaxRowNumber, level0.MaximumRow);
                Assert.IsNull(level0.Ranges);
                Assert.AreEqual(128, level0.Children.Count());
                Assert.True(level0.Children.All(child => child.Level == 1));
                Assert.AreEqual(64, level0.Children.Count(child =>
                    child.MinimumColumn == 1 &&
                    child.MaximumColumn == 8192 &&
                    child.X == 0));
                Assert.AreEqual(64, level0.Children.Count(child =>
                    child.MinimumColumn == 8193 &&
                    child.MaximumColumn == 16384 &&
                    child.X == 1));
                Assert.AreEqual(2, level0.Children.Count(child =>
                    child.MinimumRow == 1 &&
                    child.MaximumRow == 8192 &&
                    child.Y == 0));
                Assert.AreEqual(2, level0.Children.Count(child =>
                    child.MinimumRow == 16385 &&
                    child.MaximumRow == 24576 &&
                    child.Y == 2));

                Assert.True(level0.Children.ElementAt(0).Children.Any());
                Assert.True(level0.Children.Skip(1).All(child => child.Children == null));

                var level8 = level0
                    .Children.First() // 1
                    .Children.First() // 2
                    .Children.First() // 3
                    .Children.First() // 4
                    .Children.First() // 5
                    .Children.First() // 6
                    .Children.First() // 7
                    .Children.Last(); // 8

                Assert.AreEqual(65, level8.MinimumColumn);
                Assert.AreEqual(65, level8.MinimumRow);
                Assert.AreEqual(128, level8.MaximumColumn);
                Assert.AreEqual(128, level8.MaximumRow);

                var level9 = level8.Children.First();
                Assert.NotNull(level9.Ranges);
                Assert.AreEqual(range, level9.Ranges.Single());
            }
        }

        [Test]
        public void XLRangesCountChangesCorrectly()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
                var range1 = ws.Range("A1:B2");
                var range2 = ws.Range("A2:B3");
                var range3 = ws.Range("A1:B2"); // same as range1

                var ranges = new XLRanges();
                ranges.Add(range1);
                Assert.AreEqual(1, ranges.Count);
                ranges.Add(range2);
                Assert.AreEqual(2, ranges.Count);
                ranges.Add(range3);
                Assert.AreEqual(2, ranges.Count);

                Assert.AreEqual(ranges.Count, ranges.Count());

                // Add many entries to activate QuadTree
                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    ranges.Add(ws.Range(i * 2, 2, i * 2, 4));
                }

                Assert.AreEqual(2 + TEST_COUNT, ranges.Count);

                for (int i = 1; i <= TEST_COUNT; i++)
                {
                    ranges.Remove(ws.Range(i * 2, 2, i * 2, 4));
                }

                Assert.AreEqual(2, ranges.Count);

                ranges.Remove(range3);
                Assert.AreEqual(1, ranges.Count);
                ranges.Remove(range2);
                Assert.AreEqual(0, ranges.Count);
                ranges.Remove(range1);
                Assert.AreEqual(0, ranges.Count);
            }
        }

        private IXLRangeIndex CreateRangeIndex(IXLWorksheet worksheet)
        {
            return new XLRangeIndex<IXLRangeBase>((XLWorksheet)worksheet);
        }

        private IXLRangeIndex FillIndexWithTestData(IXLWorksheet worksheet)
        {
            var ranges = new List<IXLRange>();
            for (int i = 1; i <= TEST_COUNT; i++)
            {
                ranges.Add(worksheet.Range(i * 2, 2, i * 2, 4));
            }

            var index = CreateRangeIndex(worksheet);
            ranges.ForEach(r => index.Add(r));
            return index;
        }
    }
}
