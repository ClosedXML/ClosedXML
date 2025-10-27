using System;
using System.Collections.Generic;
using System.Linq;
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

            Assert.AreNotEqual(fontKey1.GetHashCode(), fontKey2.GetHashCode());
            Assert.AreEqual(fontKey2.GetHashCode(), fontKey3.GetHashCode());
        }

        [Test]
        public void XLFontKey_Equals_IsCaseInsensitive()
        {
            var fontKey1 = _defaultKey with { FontName = "Arial" };
            var fontKey2 = _defaultKey with { FontName = "Times New Roman" };
            var fontKey3 = _defaultKey with { FontName = "TIMES NEW ROMAN" };

            Assert.IsFalse(fontKey1.Equals(fontKey2));
            Assert.IsTrue(fontKey2.Equals(fontKey3));
        }

        [Test]
        [TestCaseSource(nameof(FontProperties))]
        public void Font_property_can_be_set(IFormatTestCase<IXLFont> testCase)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cellsFormat = ((XLCells)ws.Cells("A1:C4")).Format;
            var cellFormat = ((XLCell)ws.Cell("B2")).Format;

            foreach (var testValue in testCase.Values)
            {
                testCase.SetPropertyValue(cellsFormat.Font, testValue);
                var setValue = testCase.GetPropertyValue(cellFormat.Font);
                Assert.AreEqual(testValue, setValue);
            }
        }

        private static IEnumerable<object> FontProperties()
        {
            yield return new FontTestCase<bool>(font => font.Bold, (font, value) => font.Bold = value, true, false);
            yield return new FontTestCase<bool>(font => font.Bold, (font, value) => font.SetBold(value), true, false);
            yield return new FontTestCase<bool>(font => font.Bold, (font, _) => font.SetBold(), true);

            yield return new FontTestCase<bool>(font => font.Italic, (font, value) => font.Italic = value, true, false);
            yield return new FontTestCase<bool>(font => font.Italic, (font, value) => font.SetItalic(value), true, false);
            yield return new FontTestCase<bool>(font => font.Italic, (font, _) => font.SetItalic(), true);

            yield return new FontTestCase<string>(font => font.FontName, (font, value) => font.FontName = value, "Calibri", "Arial", "Consolas");
            yield return new FontTestCase<string>(font => font.FontName, (font, value) => font.SetFontName(value), "Calibri", "Arial", "Consolas");

            yield return new FontTestCase<double>(font => font.FontSize, (font, value) => font.FontSize = value, 1, 15, 409.55);
            yield return new FontTestCase<double>(font => font.FontSize, (font, value) => font.SetFontSize(value), 1, 15, 409.55);
        }

        private class FontTestCase<T> : IFormatTestCase<IXLFont>
        {
            private readonly Func<IXLFont, T> _getter;
            private readonly Action<IXLFont, T> _setter;
            private readonly IReadOnlyList<T> _testValues;

            public FontTestCase(Func<IXLFont, T> getter, Action<IXLFont, T> setter, params T[] testValues)
            {
                _getter = getter;
                _setter = setter;
                _testValues = testValues;
            }

            public IEnumerable<object> Values => _testValues.Cast<object>();

            public object GetPropertyValue(IXLFont font)
            {
                return _getter(font);
            }

            public void SetPropertyValue(IXLFont font, object value)
            {
                _setter(font, (T)value);
            }
        }
    }
}
