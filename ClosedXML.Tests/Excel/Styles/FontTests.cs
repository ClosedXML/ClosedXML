using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class FontTests
    {
        private readonly XLFontKey _defaultKey = XLFontValue.Default.Key;

        [Test]
        public void XLFontKey_GetHashCode_IsCaseInsensitive()
        {
            var fontKey1 = _defaultKey with { FontName = "Arial" };
            var fontKey2 = _defaultKey with { FontName = "Times New Roman" };
            var fontKey3 = _defaultKey with { FontName = "TIMES NEW ROMAN" };

            Assert.Multiple(() =>
            {
                Assert.That(fontKey2.GetHashCode(), Is.Not.EqualTo(fontKey1.GetHashCode()));
                Assert.That(fontKey3.GetHashCode(), Is.EqualTo(fontKey2.GetHashCode()));
            });
        }

        [Test]
        public void XLFontKey_Equals_IsCaseInsensitive()
        {
            var fontKey1 = _defaultKey with { FontName = "Arial" };
            var fontKey2 = _defaultKey with { FontName = "Times New Roman" };
            var fontKey3 = _defaultKey with { FontName = "TIMES NEW ROMAN" };

            Assert.Multiple(() =>
            {
                Assert.That(fontKey1, Is.Not.EqualTo(fontKey2));
                Assert.That(fontKey2, Is.EqualTo(fontKey3));
            });
        }
    }
}
