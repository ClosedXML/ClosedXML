using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class XLFillTests
    {
        [Test]
        public void BackgroundColorSetsPattern()
        {
            var fill = new XLFill {BackgroundColor = XLColor.Blue};
            Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);
        }

        [Test]
        public void BackgroundNoColorSetsPatternNone()
        {
            var fill = new XLFill {BackgroundColor = XLColor.NoColor};
            Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);
        }

        [Test]
        public void BackgroundPatternEqualCheck()
        {
            var fill1 = new XLFill {PatternBackgroundColor = XLColor.Blue};
            var fill2 = new XLFill {PatternBackgroundColor = XLColor.Blue};
            Assert.IsTrue(fill1.Equals(fill2));
        }

        [Test]
        public void BackgroundPatternNotEqualCheck()
        {
            var fill1 = new XLFill {PatternBackgroundColor = XLColor.Blue};
            var fill2 = new XLFill {PatternBackgroundColor = XLColor.Red};
            Assert.IsFalse(fill1.Equals(fill2));
        }

        [Test]
        public void BackgroundWithConditionalFormat()
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Test");
            worksheet.Cell(2, 2).SetValue("Text");
            var cf = worksheet.Cell(2, 2).AddConditionalFormat();
            var style = cf.WhenNotBlank();
            style
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
    }
}