using System;
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
    public class XLFillTests
    {

        [TestMethod]
        public void BackgroundPatternEqualCheck()
        {
            var fill1 = new XLFill { PatternBackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { PatternBackgroundColor = XLColor.Blue };
            Assert.IsTrue(fill1.Equals(fill2));
        }

        [TestMethod]
        public void BackgroundPatternNotEqualCheck()
        {
            var fill1 = new XLFill { PatternBackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { PatternBackgroundColor = XLColor.Red };
            Assert.IsFalse(fill1.Equals(fill2));
        }

        [TestMethod]
        public void BackgroundColorSetsPattern()
        {
            var fill = new XLFill { BackgroundColor = XLColor.Blue };
            Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);
        }

        [TestMethod]
        public void BackgroundNoColorSetsPatternNone()
        {
            var fill = new XLFill { BackgroundColor = XLColor.NoColor };
            Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);
        }
    }
}
