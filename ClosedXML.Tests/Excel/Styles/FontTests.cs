using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class FontTests
    {
        [Test]
        public void XLFontKey_GetHashCode_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKey { FontName = "Arial" };
            var fontKey2 = new XLFontKey { FontName = "Times New Roman" };
            var fontKey3 = new XLFontKey { FontName = "TIMES NEW ROMAN" };

            Assert.AreNotEqual(fontKey1.GetHashCode(), fontKey2.GetHashCode());
            Assert.AreEqual(fontKey2.GetHashCode(), fontKey3.GetHashCode());
        }

        [Test]
        public void XLFontKey_Equals_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKey { FontName = "Arial" };
            var fontKey2 = new XLFontKey { FontName = "Times New Roman" };
            var fontKey3 = new XLFontKey { FontName = "TIMES NEW ROMAN" };

            Assert.IsFalse(fontKey1.Equals(fontKey2));
            Assert.IsTrue(fontKey2.Equals(fontKey3));
        }
    }
}
