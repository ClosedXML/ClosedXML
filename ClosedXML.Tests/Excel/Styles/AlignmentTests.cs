using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    [TestFixture]
    public class AlignmentTests
    {
        [Test]
        public void TextRotationCanBeFromMinus90To90DegreesAnd255ForVerticalLayout()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();
                ws.ColumnWidth = 10;
                ws.Cell(1, 1)
                    .SetValue("Vertical: 255")
                    .Style.Alignment.SetTextRotation(255);

                for (var angle = -90; angle <= +90; angle += 10)
                {
                    var column = (angle + 90) / 10 + 2;
                    var cell = ws.Cell(1, column);
                    cell.Value = $"Rotation: {angle}";
                    cell.Style.Alignment.TextRotation = angle;
                }
            }, @"Other\Styles\Alignment\TextRotation.xlsx");
        }

        [Test]
        public void TextRotationIsConvertedOnLoadToMinus90To90Degrees()
        {
            TestHelper.LoadAndAssert(wb =>
            {
                var ws = wb.Worksheets.Single();
                Assert.AreEqual(255, ws.Cell(1,1).Style.Alignment.TextRotation);
                for (var column = 2; column < 21; ++column)
                {
                    var expectedAngle = (column - 2) * 10 - 90;
                    Assert.AreEqual(expectedAngle, ws.Cell(1, column).Style.Alignment.TextRotation);
                }
            }, @"Other\Styles\Alignment\TextRotation.xlsx");
        }

        [TestCase(91)]
        [TestCase(-91)]
        [TestCase(254)]
        [TestCase(256)]
        public void TextRotationOutsideBoundsThrowsException(int textRotation)
        {
            Assert.Throws<ArgumentException>(() =>
            {
                using var wb = new XLWorkbook();
                var ws = wb.AddWorksheet();
                ws.FirstCell().Style.Alignment.TextRotation = textRotation;
            });
        }
    }
}
