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
        public void Font_property_can_be_individually_set(IFormatTestCase<IXLFont> testCase)
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

        [Test]
        public void Font_can_be_set_by_assigning_font()
        {
            // Arrange
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Style
                .Font.SetBold()
                .Font.SetItalic()
                .Font.SetUnderline(XLFontUnderlineValues.DoubleAccounting)
                .Font.SetStrikethrough()
                .Font.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Superscript)
                .Font.SetShadow()
                .Font.SetFontSize(25)
                .Font.SetFontColor(XLColor.Red)
                .Font.SetFontName("Arial")
                .Font.SetFontFamilyNumbering(XLFontFamilyNumberingValues.Decorative)
                .Font.SetFontCharSet(XLFontCharSet.Hangul)
                .Font.SetFontScheme(XLFontScheme.Minor);

            // Act
            ws.Cell("A2").Style.Font = ws.Cell("A1").Style.Font;

            // Assert
            var copiedFont = ws.Cell("A2").Style.Font;
            Assert.IsTrue(copiedFont.Bold);
            Assert.IsTrue(copiedFont.Italic);
            Assert.AreEqual(XLFontUnderlineValues.DoubleAccounting, copiedFont.Underline);
            Assert.IsTrue(copiedFont.Strikethrough);
            Assert.AreEqual(XLFontVerticalTextAlignmentValues.Superscript, copiedFont.VerticalAlignment);
            Assert.IsTrue(copiedFont.Shadow);
            Assert.AreEqual(25, copiedFont.FontSize);
            Assert.AreEqual(XLColor.Red, copiedFont.FontColor);
            Assert.AreEqual("Arial", copiedFont.FontName);
            Assert.AreEqual(XLFontFamilyNumberingValues.Decorative, copiedFont.FontFamilyNumbering);
            Assert.AreEqual(XLFontCharSet.Hangul, copiedFont.FontCharSet);
            Assert.AreEqual(XLFontScheme.Minor, copiedFont.FontScheme);
        }

        [Test]
        public void Font_can_be_checked_for_equality()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var testFont = ws.Cell("A1").Style.Font;
            var equalFont = ws.Cell("A2").Style.Font;

            Assert.AreEqual(testFont, equalFont);
            var makeDifferentFont = new Action<IXLFont>[]
            {
                x => x.Bold = !x.Bold,
                x => x.Italic = !x.Italic,
                x => x.Underline = XLFontUnderlineValues.DoubleAccounting,
                x => x.Strikethrough = !x.Strikethrough,
                x => x.VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript,
                x => x.Shadow = !x.Shadow,
                x => x.FontSize = 25,
                x => x.FontColor = XLColor.Blue,
                x => x.FontName = "Arial",
                x => x.FontFamilyNumbering = XLFontFamilyNumberingValues.Decorative,
                x => x.FontCharSet = XLFontCharSet.Arabic,
                x => x.FontScheme = XLFontScheme.Minor,
            };
            var cell = ws.Cell("A3");
            foreach (var modify in makeDifferentFont)
            {
                modify(cell.Style.Font);
                Assert.AreNotEqual(testFont, cell.Style.Font);
                cell = cell.CellRight();
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

            var underlineValues = GetEnumValues<XLFontUnderlineValues>();
            yield return new FontTestCase<XLFontUnderlineValues>(font => font.Underline, (font, value) => font.Underline = value, underlineValues);
            yield return new FontTestCase<XLFontUnderlineValues>(font => font.Underline, (font, value) => font.SetUnderline(value), underlineValues);
            yield return new FontTestCase<XLFontUnderlineValues>(font => font.Underline, (font, _) => font.SetUnderline(), XLFontUnderlineValues.Single);

            yield return new FontTestCase<bool>(font => font.Strikethrough, (font, value) => font.Strikethrough = value, true, false);
            yield return new FontTestCase<bool>(font => font.Strikethrough, (font, value) => font.SetStrikethrough(value), true, false);
            yield return new FontTestCase<bool>(font => font.Strikethrough, (font, _) => font.SetStrikethrough(), true);

            var valignValues = GetEnumValues<XLFontVerticalTextAlignmentValues>();
            yield return new FontTestCase<XLFontVerticalTextAlignmentValues>(font => font.VerticalAlignment, (font, value) => font.VerticalAlignment = value, valignValues);
            yield return new FontTestCase<XLFontVerticalTextAlignmentValues>(font => font.VerticalAlignment, (font, value) => font.SetVerticalAlignment(value), valignValues);

            yield return new FontTestCase<bool>(font => font.Shadow, (font, value) => font.Shadow = value, true, false);
            yield return new FontTestCase<bool>(font => font.Shadow, (font, value) => font.SetShadow(value), true, false);
            yield return new FontTestCase<bool>(font => font.Shadow, (font, _) => font.SetShadow(), true);

            yield return new FontTestCase<double>(font => font.FontSize, (font, value) => font.FontSize = value, 1, 15, 409.55);
            yield return new FontTestCase<double>(font => font.FontSize, (font, value) => font.SetFontSize(value), 1, 15, 409.55);

            yield return new FontTestCase<XLColor>(font => font.FontColor, (font, value) => font.FontColor = value, XLColor.Black, XLColor.Red, XLColor.Auto);
            yield return new FontTestCase<XLColor>(font => font.FontColor, (font, value) => font.SetFontColor(value), XLColor.Black, XLColor.Red, XLColor.Auto);

            yield return new FontTestCase<string>(font => font.FontName, (font, value) => font.FontName = value, "Calibri", "Arial", "Consolas");
            yield return new FontTestCase<string>(font => font.FontName, (font, value) => font.SetFontName(value), "Calibri", "Arial", "Consolas");

            var familyValues = GetEnumValues<XLFontFamilyNumberingValues>();
            yield return new FontTestCase<XLFontFamilyNumberingValues>(font => font.FontFamilyNumbering, (font, value) => font.FontFamilyNumbering = value, familyValues);
            yield return new FontTestCase<XLFontFamilyNumberingValues>(font => font.FontFamilyNumbering, (font, value) => font.SetFontFamilyNumbering(value), familyValues);

            var charsetValues = GetEnumValues<XLFontCharSet>();
            yield return new FontTestCase<XLFontCharSet>(font => font.FontCharSet, (font, value) => font.FontCharSet = value, charsetValues);
            yield return new FontTestCase<XLFontCharSet>(font => font.FontCharSet, (font, value) => font.SetFontCharSet(value), charsetValues);

            var schemeValues = GetEnumValues<XLFontScheme>();
            yield return new FontTestCase<XLFontScheme>(font => font.FontScheme, (font, value) => font.FontScheme = value, schemeValues);
            yield return new FontTestCase<XLFontScheme>(font => font.FontScheme, (font, value) => font.SetFontScheme(value), schemeValues);
        }

        // TODO: Replace with EnumPolyfill once Polyfill is updated
        private static T[] GetEnumValues<T>()
            where T : struct, Enum
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToArray();
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
