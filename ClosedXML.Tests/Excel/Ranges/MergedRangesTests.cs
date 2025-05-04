using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests
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

            string first = ws.FirstCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();
            string last = ws.LastCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();

            Assert.Multiple(() =>
            {
                Assert.That(first, Is.EqualTo("B2"));
                Assert.That(last, Is.EqualTo("D4"));
            });
        }

        [TestCase("A1:A2", "A1:A2")]
        [TestCase("A2:B2", "A2:B2")]
        [TestCase("A3:C3", "A3:E3")]
        [TestCase("B4:B6", "B4:B6")]
        [TestCase("C7:D7", "E7:F7")]
        public void MergedRangesShiftedOnColumnInsert(string originalRange, string expectedRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            var range = ws.Range(originalRange).Merge();

            ws.Column(2).InsertColumnsAfter(2);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(mr.Single(), Is.SameAs(range));
                Assert.That(range.RangeAddress.ToString(), Is.EqualTo(expectedRange));
            });
        }

        [TestCase("A1:B1", "A1:B1")]
        [TestCase("B1:B2", "B1:B2")]
        [TestCase("C1:C3", "C1:C5")]
        [TestCase("D2:F2", "D2:F2")]
        [TestCase("G4:G5", "G6:G7")]
        public void MergedRangesShiftedOnRowInsert(string originalRange, string expectedRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            var range = ws.Range(originalRange).Merge();

            ws.Row(2).InsertRowsBelow(2);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(mr.Single(), Is.SameAs(range));
                Assert.That(range.RangeAddress.ToString(), Is.EqualTo(expectedRange));
            });
        }

        [TestCase("A1:A2", true, "A1:A2")]
        [TestCase("A2:B2", true, "A2:A2")]
        [TestCase("A3:C3", true, "A3:B3")]
        [TestCase("B4:B6", false, "")]
        [TestCase("C7:D7", true, "B7:C7")]
        public void MergedRangesShiftedOnColumnDelete(string originalRange, bool expectedExist, string expectedRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            var range = ws.Range(originalRange).Merge();

            ws.Column(2).Delete();

            var mr = ws.MergedRanges.ToArray();
            if (expectedExist)
            {
                Assert.That(mr, Has.Length.EqualTo(1));
                Assert.Multiple(() =>
                {
                    Assert.That(mr.Single(), Is.SameAs(range));
                    Assert.That(range.RangeAddress.ToString(), Is.EqualTo(expectedRange));
                });
            }
            else
            {
                Assert.Multiple(() =>
                {
                    Assert.That(mr.Length, Is.EqualTo(0));
                    Assert.That(range.RangeAddress.IsValid, Is.False);
                });
            }
        }

        [TestCase("A1:B1", true, "A1:B1")]
        [TestCase("B1:B2", true, "B1:B1")]
        [TestCase("C1:C3", true, "C1:C2")]
        [TestCase("D2:F2", false, "")]
        [TestCase("G4:G5", true, "G3:G4")]
        public void MergedRangesShiftedOnRowDelete(string originalRange, bool expectedExist, string expectedRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            var range = ws.Range(originalRange).Merge();

            ws.Row(2).Delete();

            var mr = ws.MergedRanges.ToArray();
            if (expectedExist)
            {
                Assert.That(mr, Has.Length.EqualTo(1));
                Assert.Multiple(() =>
                {
                    Assert.That(mr.Single(), Is.SameAs(range));
                    Assert.That(range.RangeAddress.ToString(), Is.EqualTo(expectedRange));
                });
            }
            else
            {
                Assert.Multiple(() =>
                {
                    Assert.That(mr.Length, Is.EqualTo(0));
                    Assert.That(range.RangeAddress.IsValid, Is.False);
                });
            }
        }

        [Test]
        public void ShiftRangeRightBreaksMerges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("B4:C5").Merge();
            ws.Range("F2:G3").Merge(); // to be broken
            ws.Range("F4:G5").Merge(); // to be broken
            ws.Range("H1:I2").Merge();
            ws.Range("H5:I6").Merge();

            ws.Range("D3:E4").InsertColumnsAfter(2);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(mr[0].RangeAddress.ToString(), Is.EqualTo("H1:I2"));
                Assert.That(mr[1].RangeAddress.ToString(), Is.EqualTo("B2:C3"));
                Assert.That(mr[2].RangeAddress.ToString(), Is.EqualTo("B4:C5"));
                Assert.That(mr[3].RangeAddress.ToString(), Is.EqualTo("H5:I6"));
            });
        }

        [Test]
        public void ShiftRangeLeftBreaksMerges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("B4:C5").Merge();
            ws.Range("F2:G3").Merge(); // to be broken
            ws.Range("F4:G5").Merge(); // to be broken
            ws.Range("H1:I2").Merge();
            ws.Range("H5:I6").Merge();

            ws.Range("D3:E4").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(mr[0].RangeAddress.ToString(), Is.EqualTo("H1:I2"));
                Assert.That(mr[1].RangeAddress.ToString(), Is.EqualTo("B2:C3"));
                Assert.That(mr[2].RangeAddress.ToString(), Is.EqualTo("B4:C5"));
                Assert.That(mr[3].RangeAddress.ToString(), Is.EqualTo("H5:I6"));
            });
        }

        [Test]
        public void RangeShiftDownBreaksMerges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("D2:E3").Merge();
            ws.Range("B6:C7").Merge(); // to be broken
            ws.Range("D6:E7").Merge(); // to be broken
            ws.Range("A8:B9").Merge();
            ws.Range("E8:F9").Merge();

            ws.Range("C4:D5").InsertRowsBelow(2);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(mr[0].RangeAddress.ToString(), Is.EqualTo("B2:C3"));
                Assert.That(mr[1].RangeAddress.ToString(), Is.EqualTo("D2:E3"));
                Assert.That(mr[2].RangeAddress.ToString(), Is.EqualTo("A8:B9"));
                Assert.That(mr[3].RangeAddress.ToString(), Is.EqualTo("E8:F9"));
            });
        }

        [Test]
        public void RangeShiftUpBreaksMerges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("MRShift");
            ws.Range("B2:C3").Merge();
            ws.Range("D2:E3").Merge();
            ws.Range("B6:C7").Merge(); // to be broken
            ws.Range("D6:E7").Merge(); // to be broken
            ws.Range("A8:B9").Merge();
            ws.Range("E8:F9").Merge();

            ws.Range("C4:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);

            var mr = ws.MergedRanges.ToArray();
            Assert.That(mr, Has.Length.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(mr[0].RangeAddress.ToString(), Is.EqualTo("B2:C3"));
                Assert.That(mr[1].RangeAddress.ToString(), Is.EqualTo("D2:E3"));
                Assert.That(mr[2].RangeAddress.ToString(), Is.EqualTo("A8:B9"));
                Assert.That(mr[3].RangeAddress.ToString(), Is.EqualTo("E8:F9"));
            });
        }

        [Test]
        public void MergedCellsAcquireFirstCellStyle()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
            ws.Range("A1:A3").Merge();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(ws.Cell("A3").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            });
        }

        [Test]
        public void MergedCellsLooseData()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A3").SetValue(100);
            ws.Range("A1:A3").Merge();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(100));
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(Blank.Value));
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void MergedCellsLooseConditionalFormats()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").AddConditionalFormat().WhenContains("1").Fill.BackgroundColor = XLColor.Red;
            ws.Cell("A2").AddConditionalFormat().WhenContains("2").Fill.BackgroundColor = XLColor.Yellow;

            ws.Range("A1:A2").Merge();

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            });
        }

        [Test]
        public void MergedCellsLooseDataValidation()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").CreateDataValidation().WholeNumber.Between(1, 2);
            ws.Cell("A2").CreateDataValidation().Date.GreaterThan(new System.DateTime(2018, 1, 1));

            ws.Range("A1:A2").Merge();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").HasDataValidation, Is.True);
                Assert.That(ws.Cell("A1").GetDataValidation().MinValue, Is.EqualTo("1"));
                Assert.That(ws.Cell("A1").GetDataValidation().MaxValue, Is.EqualTo("2"));
                Assert.That(ws.Cell("A2").HasDataValidation, Is.False);
            });
        }

        [Test]
        public void UnmergedCellsPreserveStyle()
        {
            using XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var range = ws.Range("B2:D4");
            range.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.DarkBlue)
                .Border.SetInsideBorder(XLBorderStyleValues.Thin)
                .Border.SetInsideBorderColor(XLColor.Pink);
            range.Cells().ForEach(c => c.Value = c.Address.ToString());

            var firstCell = ws.Cell("B2");
            firstCell.Style.Fill.SetBackgroundColor(XLColor.Red)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold();

            range.Merge();
            range.Unmerge();

            Assert.Multiple(() =>
            {
                Assert.That(range.Cells().All(c => c.Style.Fill.BackgroundColor == XLColor.Red), Is.True);
                Assert.That(range.Cells().Where(c => !c.Equals(firstCell)).All(c => c.Value.Equals(Blank.Value)), Is.True);
                Assert.That(firstCell.Value, Is.EqualTo("B2"));

                Assert.That(ws.Cell("B2").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("B2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B2").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thick));

                Assert.That(ws.Cell("C2").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("C2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C2").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));

                Assert.That(ws.Cell("D2").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("D2").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("D2").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("D2").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));

                Assert.That(ws.Cell("B3").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B3").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B3").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B3").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thick));

                Assert.That(ws.Cell("C3").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C3").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C3").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C3").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));

                Assert.That(ws.Cell("D3").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("D3").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("D3").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("D3").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));

                Assert.That(ws.Cell("B4").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B4").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("B4").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("B4").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thick));

                Assert.That(ws.Cell("C4").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C4").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("C4").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("C4").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));

                Assert.That(ws.Cell("D4").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
                Assert.That(ws.Cell("D4").Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("D4").Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Thick));
                Assert.That(ws.Cell("D4").Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));
            });
        }

        [Test]
        public void MergedRangesCellValuesShouldNotBeSet()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge();
                ws.Cell("A2").Value = 1;
                ws.Cell("A3").Value = 1;
                ws.Cell("A4").Value = 1;
                ws.Cell("B1").FormulaA1 = "SUM(A:A)";
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge().SetValue(1);
                ws.Cell("B1").FormulaA1 = "SUM(A:A)";
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }
        }

        [Test]
        public void MergedRangesCellFormulasShouldNotBeSet()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge();
                ws.Cell("A2").FormulaA1 = "=1";
                ws.Cell("A3").FormulaA1 = "=1";
                ws.Cell("A4").FormulaA1 = "=1";
                ws.Cell("B1").FormulaA1 = "SUM(A:A)";
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge();
                ws.Cell("A2").SetFormulaA1("=1");
                ws.Cell("A3").SetFormulaA1("=1");
                ws.Cell("A4").SetFormulaA1("=1");
                ws.Cell("B1").SetFormulaA1("SUM(A:A)");
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge();
                ws.Cell("A2").FormulaR1C1 = "=1";
                ws.Cell("A3").FormulaR1C1 = "=1";
                ws.Cell("A4").FormulaR1C1 = "=1";
                ws.Cell("B1").FormulaR1C1 = "SUM(A:A)";
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet();
                ws.Range("A2:A4").Merge();
                ws.Cell("A2").SetFormulaR1C1("=1");
                ws.Cell("A3").SetFormulaR1C1("=1");
                ws.Cell("A4").SetFormulaR1C1("=1");
                ws.Cell("B1").SetFormulaR1C1("SUM(A:A)");
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(1));
            }
        }

        [Test]
        public void MergeSingleCellRangeDoesNothing()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range(1, 1, 1, 1);

            range.Merge();

            Assert.Multiple(() =>
            {
                Assert.That(range.IsMerged(), Is.False);
                Assert.That(ws.MergedRanges.Count, Is.EqualTo(0));
            });
        }
    }
}
