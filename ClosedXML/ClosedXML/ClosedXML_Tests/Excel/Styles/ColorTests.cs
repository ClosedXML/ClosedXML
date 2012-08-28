using System;
using System.Drawing;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ColorTests
    {

        [TestMethod]
        public void ColorEqualOperatorInPlace()
        {
            Assert.IsTrue(XLColor.Black == XLColor.Black);
        }

        [TestMethod]
        public void ColorNotEqualOperatorInPlace()
        {
            Assert.IsFalse(XLColor.Black != XLColor.Black);
        }

        [TestMethod]
        public void DefaultColorIndex64isTransparentWhite()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var color = ws.FirstCell().Style.Fill.BackgroundColor;
            Assert.AreEqual(XLColorType.Indexed, color.ColorType);
            Assert.AreEqual(64, color.Indexed);
            Assert.AreEqual(Color.Transparent , color.Color);
        }
    }
}
