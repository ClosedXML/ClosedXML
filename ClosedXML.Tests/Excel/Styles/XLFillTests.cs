using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests.Excel.Styles
{
    [TestFixture]
    public class XLFillTests
    {
        [Test]
        public void BackgroundColorSetsPattern()
        {
            var fill = new XLFill { BackgroundColor = XLColor.Blue };
            Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);
        }

        [Test]
        public void BackgroundNoColorSetsPatternNone()
        {
            var fill = new XLFill { BackgroundColor = XLColor.NoColor };
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
            using var workbook = new XLWorkbook();
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
            var expectedFilePath = @"Other\StyleReferenceFiles\TransparentBackgroundFill\TransparentBackgroundFill.xlsx";

            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\TransparentBackgroundFill\inputfile.xlsx"));
            using var ms = new MemoryStream();

            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);

                wb.SaveAs(ms);

                //Uncomment to replace expectation running.net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);
                return wb;
            }, expectedFilePath);
        }
    }
}