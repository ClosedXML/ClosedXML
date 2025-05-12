using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class StylesReaderTests
{
    [Test]
    public void Can_parse_number_format()
    {
        AssertNumberFormats(
            """
            <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
            """,
            styles =>
            {
                var formatCode = styles.NumberFormats[164];
                Assert.AreEqual("\"$\"#,##0.00", formatCode);
            });
    }

    [Test]
    public void Can_read_empty_fill()
    {
        AssertFills("<fill/>", styles =>
        {
            var fill = styles.Fills.Single().Value;
            Assert.Null(fill.Pattern);
            Assert.Null(fill.LinearGradient);
            Assert.Null(fill.PathGradient);
        });
    }

    [Test]
    public void Can_read_pattern_fill()
    {
        AssertFills(
            """
            <fill>
              <patternFill patternType="lightGrid">
                <bgColor rgb="FF804000"/>
              </patternFill>
            </fill>
            """,
            styles =>
        {
            var fill = styles.Fills.Single().Value;
            Assert.NotNull(fill.Pattern);
            Assert.AreEqual(XLFillPatternValues.LightGrid, fill.Pattern.PatternType);
            Assert.AreEqual(XLColor.NoColor, fill.Pattern.PatternColor);
            Assert.AreEqual(XLColor.FromRgb(0x804000), fill.Pattern.BackgroundColor);
        });
    }

    [Test]
    public void Can_read_linear_gradient_fill()
    {
        AssertFills(
            """
            <fill>
              <gradientFill degree="90">
                <stop position="0">
                  <color rgb="FF92D050"/>
                </stop>
                <stop position="1">
                  <color rgb="FF0070C0"/>
                </stop>
              </gradientFill>
            </fill>
            """,
            styles =>
            {
                var linearGradient = styles.Fills.Single().Value.LinearGradient;
                Assert.NotNull(linearGradient);
                Assert.AreEqual(90, linearGradient.Degrees);
                Assert.That(linearGradient.Stops, Is.EquivalentTo(new Dictionary<FractionOfOne, XLColor>
                {
                    { 0, XLColor.FromRgb(0x92D050) },
                    { 1, XLColor.FromRgb(0x0070C0) }
                }));
            });
    }

    [Test]
    public void Can_read_path_gradient_fill()
    {
        AssertFills(
            """
            <fill>
              <gradientFill type="path" left="0.5" right="0.25" top="0.125" bottom="0.75">
                <stop position="0">
                  <color theme="0"/>
                </stop>
                <stop position="1">
                  <color theme="4"/>
                </stop>
              </gradientFill>
            </fill>
            """,
            styles =>
            {
                var pathGradient = styles.Fills.Single().Value.PathGradient;
                Assert.NotNull(pathGradient);
                Assert.AreEqual(0.5, pathGradient.InnerLeft);
                Assert.AreEqual(0.25, pathGradient.InnerRight);
                Assert.AreEqual(0.125, pathGradient.InnerTop);
                Assert.AreEqual(0.75, pathGradient.InnerBottom);
                Assert.That(pathGradient.Stops, Is.EquivalentTo(new Dictionary<FractionOfOne, XLColor>
                {
                    { 0, XLColor.FromTheme(XLThemeColor.Background1) },
                    { 1, XLColor.FromTheme(XLThemeColor.Accent1) },
                }));
            });
    }

    [Test]
    public void Can_read_alignment()
    {
        AssertCellXfs(
            """
            <alignment horizontal="center"
                       vertical="top"
                       textRotation="45"
                       wrapText="1"
                       indent="7"
                       relativeIndent="4"
                       justifyLastLine="1"
                       shrinkToFit="1"
                       readingOrder="2"
                       />
            """,
            styles =>
            {
                var alignment = styles.CellFormats[0].Alignment;
                Assert.NotNull(alignment);
                Assert.AreEqual(XLAlignmentHorizontalValues.Center, alignment.Horizontal);
                Assert.AreEqual(XLAlignmentVerticalValues.Top, alignment.Vertical);
                Assert.AreEqual(45, alignment.TextRotation.Value);
                Assert.IsTrue(alignment.WrapText);
                Assert.AreEqual(7, alignment.Indent);
                Assert.AreEqual(4, alignment.RelativeIndent);
                Assert.IsTrue(alignment.JustifyLastLine);
                Assert.IsTrue(alignment.ShrinkToFit);
                Assert.AreEqual(XLAlignmentReadingOrderValues.RightToLeft, alignment.ReadingOrder);
            });
    }

    [Test]
    public void Can_read_protection()
    {
        AssertCellXfs(
            """
            <protection locked="false" hidden="1"/>
            """,
            styles =>
            {
                var protection = styles.CellFormats[0].Protection;
                Assert.NotNull(protection);
                Assert.IsFalse(protection.Locked);
                Assert.IsTrue(protection.Hidden);
            });
    }

    private static void AssertNumberFormats(string numberFormatsXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     <numFmts>
                       {numberFormatsXml}
                     </numFmts>
                   </styleSheet>
                   """;
        AssertFormat(assert, xml);
    }


    private static void AssertFills(string fillsXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     <fills>
                       {fillsXml}
                     </fills>
                   </styleSheet>
                   """;
        AssertFormat(assert, xml);
    }

    private static void AssertCellXfs(string cellXfsXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     <cellXfs>
                       <xf>
                         {cellXfsXml}
                       </xf>
                     </cellXfs>
                   </styleSheet>
                   """;
        AssertFormat(assert, xml);
    }

    private static void AssertFormat(Action<XLWorkbookStyles> assert, string xml)
    {
        using var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(xml));
        using var xmlTreeReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, false);
        var styles = new XLWorkbookStyles();
        var reader = new StylesReader(xmlTreeReader, styles);
        reader.Load();
        assert(styles);
    }
}
