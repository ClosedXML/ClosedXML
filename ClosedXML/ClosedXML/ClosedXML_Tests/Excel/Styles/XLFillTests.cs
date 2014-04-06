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
    }
}