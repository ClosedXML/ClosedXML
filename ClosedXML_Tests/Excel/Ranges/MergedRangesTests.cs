using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class MergedRangesTests
    {
        [Test]
        public void LastCellFromMerge()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");
            ws.Range("B2:D4").Merge();

            string first = ws.FirstCellUsed(true).Address.ToStringRelative();
            string last = ws.LastCellUsed(true).Address.ToStringRelative();

            Assert.AreEqual("B2", first);
            Assert.AreEqual("D4", last);
        }

        [TestCase("A1:A2", "A1:A2")]
        [TestCase("A2:B2", "A2:B2")]
        [TestCase("A3:C3", "A3:E3")]
        [TestCase("B4:B6", "B4:B6")]
        [TestCase("C7:D7", "E7:F7")]
        public void MergedRangesShiftedOnColumnInsert(string originalRange, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                var range = ws.Range(originalRange).Merge();

                ws.Column(2).InsertColumnsAfter(2);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(1, mr.Length);
                Assert.AreSame(range, mr.Single());
                Assert.AreEqual(expectedRange, range.RangeAddress.ToString());
            }
        }

        [TestCase("A1:B1", "A1:B1")]
        [TestCase("B1:B2", "B1:B2")]
        [TestCase("C1:C3", "C1:C5")]
        [TestCase("D2:F2", "D2:F2")]
        [TestCase("G4:G5", "G6:G7")]
        public void MergedRangesShiftedOnRowInsert(string originalRange, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                var range = ws.Range(originalRange).Merge();

                ws.Row(2).InsertRowsBelow(2);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(1, mr.Length);
                Assert.AreSame(range, mr.Single());
                Assert.AreEqual(expectedRange, range.RangeAddress.ToString());
            }
        }

        [TestCase("A1:A2", true, "A1:A2")]
        [TestCase("A2:B2", true, "A2:A2")]
        [TestCase("A3:C3", true, "A3:B3")]
        [TestCase("B4:B6", false, "")]
        [TestCase("C7:D7", true, "B7:C7")]
        public void MergedRangesShiftedOnColumnDelete(string originalRange, bool expectedExist, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                var range = ws.Range(originalRange).Merge();

                ws.Column(2).Delete();

                var mr = ws.MergedRanges.ToArray();
                if (expectedExist)
                {
                    Assert.AreEqual(1, mr.Length);
                    Assert.AreSame(range, mr.Single());
                    Assert.AreEqual(expectedRange, range.RangeAddress.ToString());
                }
                else
                {
                    Assert.AreEqual(0, mr.Length);
                    Assert.IsFalse(range.RangeAddress.IsValid);
                }
            }
        }

        [TestCase("A1:B1", true, "A1:B1")]
        [TestCase("B1:B2", true, "B1:B1")]
        [TestCase("C1:C3", true, "C1:C2")]
        [TestCase("D2:F2", false, "")]
        [TestCase("G4:G5", true, "G3:G4")]
        public void MergedRangesShiftedOnRowDelete(string originalRange, bool expectedExist, string expectedRange)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                var range = ws.Range(originalRange).Merge();

                ws.Row(2).Delete();

                var mr = ws.MergedRanges.ToArray();
                if (expectedExist)
                {
                    Assert.AreEqual(1, mr.Length);
                    Assert.AreSame(range, mr.Single());
                    Assert.AreEqual(expectedRange, range.RangeAddress.ToString());
                }
                else
                {
                    Assert.AreEqual(0, mr.Length);
                    Assert.IsFalse(range.RangeAddress.IsValid);
                }
            }
        }

        [Test]
        public void ShiftRangeRightBreaksMerges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                ws.Range("B2:C3").Merge();
                ws.Range("B4:C5").Merge();
                ws.Range("F2:G3").Merge(); // to be broken
                ws.Range("F4:G5").Merge(); // to be broken
                ws.Range("H1:I2").Merge();
                ws.Range("H5:I6").Merge();

                ws.Range("D3:E4").InsertColumnsAfter(2);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(4, mr.Length);
                Assert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
                Assert.AreEqual("B4:C5", mr[1].RangeAddress.ToString());
                Assert.AreEqual("H1:I2", mr[2].RangeAddress.ToString());
                Assert.AreEqual("H5:I6", mr[3].RangeAddress.ToString());
            }
        }

        [Test]
        public void ShiftRangeLeftBreaksMerges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                ws.Range("B2:C3").Merge();
                ws.Range("B4:C5").Merge();
                ws.Range("F2:G3").Merge(); // to be broken
                ws.Range("F4:G5").Merge(); // to be broken
                ws.Range("H1:I2").Merge();
                ws.Range("H5:I6").Merge();

                ws.Range("D3:E4").Delete(XLShiftDeletedCells.ShiftCellsLeft);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(4, mr.Length);
                Assert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
                Assert.AreEqual("B4:C5", mr[1].RangeAddress.ToString());
                Assert.AreEqual("H1:I2", mr[2].RangeAddress.ToString());
                Assert.AreEqual("H5:I6", mr[3].RangeAddress.ToString());
            }
        }

        [Test]
        public void RangeShiftDownBreaksMerges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                ws.Range("B2:C3").Merge();
                ws.Range("D2:E3").Merge();
                ws.Range("B6:C7").Merge(); // to be broken
                ws.Range("D6:E7").Merge(); // to be broken
                ws.Range("A8:B9").Merge();
                ws.Range("E8:F9").Merge();

                ws.Range("C4:D5").InsertRowsBelow(2);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(4, mr.Length);
                Assert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
                Assert.AreEqual("D2:E3", mr[1].RangeAddress.ToString());
                Assert.AreEqual("A8:B9", mr[2].RangeAddress.ToString());
                Assert.AreEqual("E8:F9", mr[3].RangeAddress.ToString());
            }
        }

        [Test]
        public void RangeShiftUpBreaksMerges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("MRShift");
                ws.Range("B2:C3").Merge();
                ws.Range("D2:E3").Merge();
                ws.Range("B6:C7").Merge(); // to be broken
                ws.Range("D6:E7").Merge(); // to be broken
                ws.Range("A8:B9").Merge();
                ws.Range("E8:F9").Merge();

                ws.Range("C4:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);

                var mr = ws.MergedRanges.ToArray();
                Assert.AreEqual(4, mr.Length);
                Assert.AreEqual("B2:C3", mr[0].RangeAddress.ToString());
                Assert.AreEqual("D2:E3", mr[1].RangeAddress.ToString());
                Assert.AreEqual("A8:B9", mr[2].RangeAddress.ToString());
                Assert.AreEqual("E8:F9", mr[3].RangeAddress.ToString());
            }
        }
    }
}
