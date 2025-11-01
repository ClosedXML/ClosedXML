using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class XLFillTests
    {
        [Test]
        public void BackgroundColor_keeps_pattern_on_two_color_patterns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var fill = ws.Cell("A1").Style.Fill;
            fill.PatternType = XLFillPatternValues.LightGrid;
            Assert.AreEqual(XLFillPatternValues.LightGrid, fill.PatternType);

            fill.BackgroundColor = XLColor.Blue;

            Assert.AreEqual(XLFillPatternValues.LightGrid, fill.PatternType);
        }

        [Test]
        public void BackgroundColor_sets_pattern_to_solid_on_pattern_none()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var fill = ws.Cell("A1").Style.Fill;
            Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);

            fill.BackgroundColor = XLColor.Blue;

            Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);
        }

        [Test]
        public void BackgroundColor_set_to_transparent_color_sets_pattern_to_none()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var fill = ws.Cell("A1").Style.Fill;
            fill.BackgroundColor = XLColor.Red;
            Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);

            fill.BackgroundColor = XLColor.Auto;

            Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);
        }

        [Test]
        public void BackgroundPatternEqualCheck()
        {
            var fill1 = new XLFill { BackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { BackgroundColor = XLColor.Blue };
            Assert.IsTrue(fill1.Equals(fill2));
            Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
        }

        [Test]
        public void BackgroundPatternNotEqualCheck()
        {
            var fill1 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Red };
            Assert.IsFalse(fill1.Equals(fill2));
            Assert.AreNotEqual(fill1.GetHashCode(), fill2.GetHashCode());
        }

        [Test]
        public void FillsWithTransparentColorEqual()
        {
            var fill1 = new XLFill { BackgroundColor = XLColor.ElectricUltramarine, PatternType = XLFillPatternValues.None };
            var fill2 = new XLFill { BackgroundColor = XLColor.EtonBlue, PatternType = XLFillPatternValues.None };
            var fill3 = new XLFill { BackgroundColor = XLColor.FromIndex(64) };
            var fill4 = new XLFill { BackgroundColor = XLColor.NoColor };

            Assert.IsTrue(fill1.Equals(fill2));
            Assert.IsTrue(fill1.Equals(fill3));
            Assert.IsTrue(fill1.Equals(fill4));
            Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
            Assert.AreEqual(fill1.GetHashCode(), fill3.GetHashCode());
            Assert.AreEqual(fill1.GetHashCode(), fill4.GetHashCode());
        }

        [Test]
        public void SolidFillsWithDifferentPatternColorEqual()
        {
            var fill1 = new XLFill
            {
                PatternType = XLFillPatternValues.Solid,
                BackgroundColor = XLColor.Red,
                PatternColor = XLColor.Blue
            };

            var fill2 = new XLFill
            {
                PatternType = XLFillPatternValues.Solid,
                BackgroundColor = XLColor.Red,
                PatternColor = XLColor.Green
            };

            Assert.IsTrue(fill1.Equals(fill2));
            Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
        }

        [Test]
        public void BackgroundWithConditionalFormat()
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Test");
            worksheet.Cell(2, 2).SetValue("Text");
            var cf = worksheet.Cell(2, 2).AddConditionalFormat();
            var style = cf.WhenNotBlank();
            style = style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            Assert.AreEqual(style.Border.BottomBorder, XLBorderStyleValues.Thick);
            Assert.AreEqual(style.Border.TopBorder, XLBorderStyleValues.Thick);
            Assert.AreEqual(style.Border.LeftBorder, XLBorderStyleValues.Thick);
            Assert.AreEqual(style.Border.RightBorder, XLBorderStyleValues.Thick);

            Assert.AreEqual(style.Border.BottomBorderColor, XLColor.Blue);
            Assert.AreEqual(style.Border.TopBorderColor, XLColor.Blue);
            Assert.AreEqual(style.Border.LeftBorderColor, XLColor.Blue);
            Assert.AreEqual(style.Border.RightBorderColor, XLColor.Blue);
        }

        [Test]
        public void LoadAndSaveTransparentBackgroundFill()
        {
            TestHelper.LoadSaveAndCompare(
                @"Other\StyleReferenceFiles\TransparentBackgroundFill\inputfile.xlsx",
                @"Other\StyleReferenceFiles\TransparentBackgroundFill\TransparentBackgroundFill.xlsx");
        }

        [Test]
        public void ReservedFills_ReplaceWithPredefinedValues()
        {
            // If attribute or whole predefined fill is missing from the file, save predefined values
            TestHelper.LoadSaveAndCompare(
                @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Input.xlsx",
                @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Output.xlsx");
        }

        [Test]
        public void ReservedFills_MoveFillsFromReservedPositions()
        {
            // If the input doesn't have expected fill values at the reserved position s0 and 1 (can only happen
            // for non-excel sources, excel always has correct values), put expected fill at 0 and 1, but save original
            // fills to different positions if they are used.
            TestHelper.LoadSaveAndCompare(
                @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Input.xlsx",
                @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Output.xlsx");
        }
    }
}
