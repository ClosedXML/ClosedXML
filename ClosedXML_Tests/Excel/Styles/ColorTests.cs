using System.Drawing;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class ColorTests
    {
        [Test]
        public void ColorEqualOperatorInPlace()
        {
            Assert.IsTrue(XLColor.Black == XLColor.Black);
        }

        [Test]
        public void ColorNotEqualOperatorInPlace()
        {
            Assert.IsFalse(XLColor.Black != XLColor.Black);
        }

        [Test]
        public void ColorNamedVsHTML()
        {
            Assert.IsTrue(XLColor.Black == XLColor.FromHtml("#000000"));
        }

        [Test]
        public void DefaultColorIndex64isTransparentWhite()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            XLColor color = ws.FirstCell().Style.Fill.BackgroundColor;
            Assert.AreEqual(XLColorType.Indexed, color.ColorType);
            Assert.AreEqual(64, color.Indexed);
            Assert.AreEqual(Color.Transparent, color.Color);
        }
    }
}