using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.Formatting;
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
    public void Can_read_empty_font()
    {
        // Empty font is valid, it will just inherit everything
        AssertFonts("<font/>", styles =>
        {
            var font = styles.Fonts.Single().Value;
            Assert.Null(font.Name);
            Assert.Null(font.Charset);
            Assert.Null(font.Family);
            Assert.Null(font.Bold);
            Assert.Null(font.Italic);
            Assert.Null(font.Strikethrough);
            Assert.Null(font.Outline);
            Assert.Null(font.Shadow);
            Assert.Null(font.Condense);
            Assert.Null(font.Extend);
            Assert.Null(font.Color);
            Assert.Null(font.Size);
            Assert.Null(font.Underline);
            Assert.Null(font.VerticalAlignment);
            Assert.Null(font.Scheme);
        });
    }

    [Test]
    public void Can_read_font()
    {
        AssertFonts(
            """
            <font>
              <b/>
              <i/>
              <strike/>
              <condense/>
              <extend/>
              <outline/>
              <shadow/>
              <u val="double"/>
              <vertAlign val="superscript"/>
              <sz val="8.5"/>
              <color rgb="FF802010"/>
              <name val="Calibri"/>
              <family val="2"/>
              <charset val="128"/>
              <scheme val="none"/>
            </font>
            """, styles =>
        {
            var font = styles.Fonts.Single().Value;
            Assert.AreEqual("Calibri", font.Name);
            Assert.AreEqual(XLFontCharSet.ShiftJIS, font.Charset);
            Assert.AreEqual(XLFontFamilyNumberingValues.Swiss, font.Family);
            Assert.IsTrue(font.Bold);
            Assert.IsTrue(font.Italic);
            Assert.IsTrue(font.Strikethrough);
            Assert.IsTrue(font.Outline);
            Assert.IsTrue(font.Shadow);
            Assert.IsTrue(font.Condense);
            Assert.IsTrue(font.Extend);
            Assert.AreEqual(XLColor.FromRgb(0x802010), font.Color);
            Assert.AreEqual(8.5, font.Size);
            Assert.AreEqual(XLFontUnderlineValues.Double, font.Underline);
            Assert.AreEqual(XLFontVerticalTextAlignmentValues.Superscript, font.VerticalAlignment);
            Assert.AreEqual(XLFontScheme.None, font.Scheme);
        });
    }

    [TestCase(6)]
    [TestCase(14)]
    public void Interprets_undefined_font_family_values_as_unknown_font_family(int fontFamily)
    {
        // Deal with serious difference between standard and Excel. Standard only defines range of
        // numerical values, but there is no meaning assigned. Thus it makes sense to take font
        // family values allowed by standard (that have no defined meaning) and convert them
        // to unknown font family.
        AssertFonts(
            $"""
            <font>
              <family val="{fontFamily}"/>
            </font>
            """, styles =>
            {
                var font = styles.Fonts.Single().Value;
                Assert.AreEqual(XLFontFamilyNumberingValues.NotApplicable, font.Family);
            });
    }

    [Test]
    public void Can_repeat_and_reorder_font_properties()
    {
        // Excel requires basically a sequence, but spec allows to repeat properties and mix the order.
        AssertFonts(
            """
            <font>
              <name val="First Font"/>
              <name val="Second Font"/>
              <b/>
              <b val="0"/>
            </font>
            """, styles =>
            {
                var font = styles.Fonts.Single().Value;
                Assert.AreEqual("Second Font", font.Name);
                Assert.IsFalse(font.Bold);
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
    public void Can_read_cell_format_alignment()
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
    public void Can_read_cell_format_protection()
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

    [Test]
    public void Can_read_cell_format_properties()
    {
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <numFmts>
                <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
              </numFmts>
              <fonts>
                <font><color theme="1"/></font>
                <font><sz val="11"/></font>
              </fonts>
              <fills>
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="gray125"/></fill>
              </fills>
              <borders>
                <border><bottom style="double"><color indexed="64"/></bottom></border>
              </borders>
              <cellStyleXfs>
                <xf />
              </cellStyleXfs>
              <cellXfs>
                <xf numFmtId="164" fontId="1" fillId="1" borderId="0" xfId="0" quotePrefix="1" pivotButton="1"
                    applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0">
                  <alignment horizontal="left"/>
                  <protection/>
                </xf>
              </cellXfs>
              <cellStyles>
                <cellStyle name="Test" xfId="0" />
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            var format = styles.CellFormats[0];
            Assert.NotNull(format.Alignment);
            Assert.AreEqual(XLAlignmentHorizontalValues.Left, format.Alignment.Horizontal);

            Assert.NotNull(format.Protection);
            Assert.IsTrue(format.Protection.Locked);
            Assert.IsFalse(format.Protection.Hidden);

            Assert.AreSame(styles.NumberFormats[164], format.NumberFormat);
            Assert.AreSame(styles.Fonts[1], format.Font);
            Assert.AreSame(styles.Fills[1], format.Fill);
            Assert.AreSame(styles.Borders[0], format.Border);

            // All apply* are 0 or default -> nothing should be overwritten
            Assert.AreEqual(CellFormatComponents.None, format.StyleComponents);
            Assert.NotNull(format.CellStyle);
            Assert.AreEqual("Test", format.CellStyle.Name);
        }, xml);
    }

    [Test]
    public void Can_read_cell_styles()
    {
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <numFmts>
                <numFmt numFmtId="190" formatCode="0.00"/>
              </numFmts>
              <fonts>
                <font/>
                <font><sz val="15"/></font>
              </fonts>
              <fills>
                <fill><patternFill patternType="gray125"/></fill>
                <fill><patternFill patternType="lightGrid"/></fill>
              </fills>
              <borders>
                <border><bottom style="double"><color rgb="FF801040"/></bottom></border>
              </borders>
              <cellStyleXfs>
                <xf numFmtId="190" fontId="1" fillId="1" borderId="0" xfId="0" 
                    applyNumberFormat="1" applyBorder="1" applyAlignment="1" applyProtection="1">
                  <alignment horizontal="right"/>
                  <protection/>
                </xf>
              </cellStyleXfs>
              <cellStyles>
                <cellStyle name="Test" xfId="0" builtinId="26"/>
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            var style = styles.CellStyles[0];
            Assert.AreEqual("Test", style.Name);
            Assert.AreEqual(BuiltInStyleValues.Good, style.BuiltInStyle);
            Assert.IsFalse(style.Hidden);

            Assert.NotNull(style.Alignment);
            Assert.AreEqual(XLAlignmentHorizontalValues.Right, style.Alignment.Horizontal);

            Assert.NotNull(style.Protection);
            Assert.IsTrue(style.Protection.Locked);
            Assert.IsFalse(style.Protection.Hidden);

            Assert.AreSame(styles.NumberFormats[190], style.NumberFormat);
            Assert.AreEqual("0.00", style.NumberFormat);

            Assert.AreSame(styles.Fonts[1], style.Font);
            Assert.AreEqual(15.0, style.Font?.Size);

            Assert.AreSame(styles.Fills[1], style.Fill);
            Assert.AreEqual(XLFillPatternValues.LightGrid, style.Fill?.Pattern?.PatternType);

            Assert.AreSame(styles.Borders[0], style.Border);
            Assert.AreEqual(XLBorderStyleValues.Double, style.Border?.Bottom?.Style);

            // All apply* are true or default (true) -> everything should be overwritten
            Assert.AreEqual(CellFormatComponents.All, style.ApplyComponents);
        }, xml);
    }

    [Test]
    public void Outline_cell_styles_are_normalized()
    {
        // builtinId 1 (RowLevel_*) and 2 (ColLevel_*) combined with iLevel are converted
        // to an item within all builtin styles instead of requiring two combined attributes.
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <fonts><font/></fonts>
              <fills><fill/></fills>
              <borders><border/></borders>
              <cellStyleXfs>
                <xf/>
              </cellStyleXfs>
              <cellStyles>
                <cellStyle name="RowLevel_3" xfId="0" builtinId="1" iLevel="2"/>
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            var style = styles.CellStyles[0];
            Assert.AreEqual("RowLevel_3", style.Name);
            Assert.AreEqual(BuiltInStyleValues.RowLevel3, style.BuiltInStyle);
        }, xml);
    }

    [Test]
    public void Cell_styles_without_name_have_a_generated_one()
    {
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <fonts><font/></fonts>
              <fills><fill/></fills>
              <borders><border/></borders>
              <cellStyleXfs>
                <xf/>
                <xf/>
                <xf/>
              </cellStyleXfs>
              <cellStyles>
                <cellStyle name="Style 1" xfId="0"/>
                <cellStyle xfId="1"/>
                <cellStyle xfId="2"/>
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            Assert.AreEqual(3, styles.CellStyles.Count);
            Assert.AreEqual("Style 1", styles.CellStyles[0].Name);
            Assert.AreEqual("Style 2", styles.CellStyles[1].Name);
            Assert.AreEqual("Style 3", styles.CellStyles[2].Name);
        }, xml);
    }

    [Test]
    public void Can_read_cell_styles_referencing_same_cell_style_format()
    {
        // This is basically pointless, but valid situation. Cell formats reference
        // the cellStyleXfs, but there are multiple cellStyles for the cellStyleXfs.
        // Having two names for essentially one style is pretty invalid and OI-29500
        // even forbids it, but Excel reads it and there are many producers in the world.
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <fonts><font><sz val="15"/></font></fonts>
              <fills><fill/></fills>
              <borders><border/></borders>
              <cellStyleXfs>
                <xf fontId="0"/>
              </cellStyleXfs>
              <cellXfs>
                <xf xfId="0"/>
              </cellXfs>
              <cellStyles>
                <cellStyle name="Style 1" xfId="0"/>
                <cellStyle name="Style 2" xfId="0"/>
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            Assert.AreEqual(2, styles.CellStyles.Count);
            Assert.AreEqual("Style 1", styles.CellStyles[0].Name);
            Assert.AreEqual(15.0, styles.CellStyles[0].Font?.Size);

            Assert.AreEqual("Style 2", styles.CellStyles[1].Name);
            Assert.AreEqual(15.0, styles.CellStyles[1].Font?.Size);

            Assert.AreEqual(1, styles.CellFormats.Count);
            Assert.AreSame(styles.CellStyles[0], styles.CellFormats[0].CellStyle);
        }, xml);
    }

    [Test]
    public void Cell_style_formatting_records_without_cell_style_have_generated_one()
    {
        var xml =
            """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <fonts>
                <font><sz val="15"/></font>
                <font><sz val="10"/></font>
              </fonts>
              <fills><fill/></fills>
              <borders><border/></borders>
              <cellStyleXfs>
                <xf fontId="0"/>
                <xf fontId="1"/>
                <xf/>
              </cellStyleXfs>
              <cellStyles>
                <cellStyle xfId="0"/>
                <cellStyle name="Style 2" xfId="1"/>
              </cellStyles>
            </styleSheet>
            """;
        AssertFormat(styles =>
        {
            Assert.AreEqual(3, styles.CellStyles.Count);
            Assert.AreEqual("Style 1", styles.CellStyles[0].Name);
            Assert.AreEqual(15.0, styles.CellStyles[0].Font?.Size);

            Assert.AreEqual("Style 2", styles.CellStyles[1].Name);
            Assert.AreEqual(10.0, styles.CellStyles[1].Font?.Size);

            // Created style from last formatting record without a cellStyle element
            Assert.AreEqual("Style 3", styles.CellStyles[2].Name);
            Assert.IsNull(styles.CellStyles[2].Font);
        }, xml);
    }

    [Test]
    public void Can_read_differential_formats()
    {
        AssertDxf(
            """
            <dxf>
              <font>
                <b/>
              </font>
              <numFmt numFmtId="176" formatCode="0.00"/>
              <fill>
                <patternFill patternType="lightGrid">
                  <fgColor rgb="FF0000FF"/>
                  <bgColor rgb="FF00FF00"/>
                </patternFill>
              </fill>
              <border>
                <right style="thin">
                  <color rgb="FF00FF00"/>
                </right>
              </border>
            </dxf>
            """,
            styles =>
            {
                var (dxfId, dxf) = styles.DifferentialFormats.Single();
                Assert.AreEqual(0, dxfId);
                Assert.IsTrue(dxf.Font?.Bold);
                Assert.AreEqual("0.00", dxf.NumberFormat);
                Assert.AreEqual(XLFillPatternValues.LightGrid, dxf.Fill?.Pattern?.PatternType);
                Assert.AreEqual(XLColor.FromRgb(0x0000FF), dxf.Fill.Pattern.PatternColor);
                Assert.AreEqual(XLColor.FromRgb(0x00FF00), dxf.Fill.Pattern.BackgroundColor);
                Assert.AreEqual(XLBorderStyleValues.Thin, dxf.Border?.Right?.Style);
                Assert.AreEqual(XLColor.FromRgb(0x00FF00), dxf.Border?.Right?.Color);
                Assert.IsNull(dxf.Border.Left);
                Assert.IsNull(dxf.Border.Top);
                Assert.IsNull(dxf.Border.Bottom);
            });
    }

    [Test]
    public void Differential_formats_do_not_have_to_specify_any_component()
    {
        AssertDxf(
            "<dxf/>",
            styles =>
            {
                var dxf = styles.DifferentialFormats.Single().Value;
                Assert.IsNull(dxf.Font);
                Assert.IsNull(dxf.NumberFormat);
                Assert.IsNull(dxf.Fill);
                Assert.IsNull(dxf.Border);
            });
    }

    [Test]
    public void Default_pattern_type_for_differential_formats_is_solid()
    {
        AssertDxf(
            """
            <dxf>
              <fill>
                <patternFill/>
              </fill>
            </dxf>
            """,
            styles =>
            {
                var dxf = styles.DifferentialFormats.Single().Value;
                Assert.AreEqual(XLFillPatternValues.Solid, dxf.Fill?.Pattern?.PatternType);
            });
    }

    [Test]
    public void Differential_formats_use_background_color_for_solid_fill_color()
    {
        AssertDxf(
            """
            <dxf>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FF00FF00"/>
                  <bgColor rgb="FF800000"/>
                </patternFill>
              </fill>
            </dxf>
            """,
            styles =>
            {
                var dxf = styles.DifferentialFormats.Single().Value;
                Assert.AreEqual(XLFillPatternValues.Solid, dxf.Fill?.Pattern?.PatternType);
                Assert.AreEqual(XLColor.FromRgb(0x800000), dxf.Fill?.Pattern?.PatternColor);
                Assert.AreEqual(XLColor.FromRgb(0x00FF00), dxf.Fill?.Pattern?.BackgroundColor);
            });
    }

    [TestCase(true, true)]
    [TestCase(true, false)]
    [TestCase(false, true)]
    [TestCase(false, false)]
    public void Only_reads_table_or_pivot_style_when_flag_is_set(bool table, bool pivot)
    {
        var xml =
            $"""
            <tableStyles>
              <tableStyle name="Test Style" table="{(table ? 1 : 0)}" pivot="{(pivot ? 1 : 0)}"/>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            Assert.AreEqual(table, styles.TableStyles.ContainsKey("Test Style"));
            Assert.AreEqual(pivot, styles.PivotStyles.ContainsKey("Test Style"));
        });
    }

    [Test]
    public void Can_read_table_style()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
              <dxf><font><sz val="10"/></font></dxf>
              <dxf><font><color rgb="FF00FF00"/></font></dxf>
              <dxf><font><color rgb="FF0000FF"/></font></dxf>
            </dxfs>
            <tableStyles count="1"
                         defaultTableStyle="TableStyleMedium2">
              <tableStyle name="Test Style"
                          pivot="0"
                          count="6">
                <tableStyleElement type="wholeTable" dxfId="0"/>
                <tableStyleElement type="headerRow" dxfId="1"/>
                <tableStyleElement type="firstRowStripe" size="2" dxfId="0"/>
                <tableStyleElement type="secondRowStripe" size="3" dxfId="1"/>
                <tableStyleElement type="firstColumnStripe" size="4" dxfId="2"/>
                <tableStyleElement type="secondColumnStripe" size="5" dxfId="3"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            Assert.AreEqual("TableStyleMedium2", styles.DefaultTableStyle);

            // Style names are case insensitive
            var tableStyle = styles.TableStyles["test style"];
            Assert.AreEqual("Test Style", tableStyle.Name);

            Assert.That(tableStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLTableStyleRegionValues, XLDifferentialFormat>
            {
                { XLTableStyleRegionValues.WholeTable, styles.DifferentialFormats[0] },
                { XLTableStyleRegionValues.HeaderRow, styles.DifferentialFormats[1] },
                { XLTableStyleRegionValues.FirstRowStripe, styles.DifferentialFormats[0] },
                { XLTableStyleRegionValues.SecondRowStripe, styles.DifferentialFormats[1] },
                { XLTableStyleRegionValues.FirstColumnStripe, styles.DifferentialFormats[2] },
                { XLTableStyleRegionValues.SecondColumnStripe, styles.DifferentialFormats[3] },
            }));

            Assert.AreEqual(2, tableStyle.RowStripe1BandSize);
            Assert.AreEqual(3, tableStyle.RowStripe2BandSize);
            Assert.AreEqual(4, tableStyle.ColumnStripe1BandSize);
            Assert.AreEqual(5, tableStyle.ColumnStripe2BandSize);
        });
    }

    [Test]
    public void Can_read_table_style_with_repeated_elements()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
              <dxf><font><color rgb="FF00FF00"/></font></dxf>
            </dxfs>
            <tableStyles>
              <tableStyle name="Test Style">
                <tableStyleElement type="wholeTable" dxfId="0"/>
                <tableStyleElement type="wholeTable" dxfId="1"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            // Take last element
            var tableStyle = styles.TableStyles["Test Style"];
            Assert.That(tableStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLTableStyleRegionValues, XLDifferentialFormat>
            {
                { XLTableStyleRegionValues.WholeTable, styles.DifferentialFormats[1] }
            }));
        });
    }

    [Test]
    public void Ignores_table_style_elements_that_are_only_for_pivot_table()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
            </dxfs>
            <tableStyles>
              <tableStyle name="Test Style">
                <tableStyleElement type="firstSubtotalColumn" dxfId="0"/>
                <tableStyleElement type="headerRow" dxfId="0"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            var tableStyle = styles.TableStyles["Test Style"];
            Assert.That(tableStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLTableStyleRegionValues, XLDifferentialFormat>
            {
                { XLTableStyleRegionValues.HeaderRow, styles.DifferentialFormats[0] }
            }));
        });
    }

    [Test]
    public void Ignores_table_style_elements_without_differential_format()
    {
        var xml =
            """
            <tableStyles>
              <tableStyle name="Test Style">
                <tableStyleElement type="totalRow" />
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            var tableStyle = styles.TableStyles["Test Style"];
            Assert.That(tableStyle.RegionFormats, Is.Empty);
        });
    }

    [Test]
    public void Can_read_pivot_style()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
              <dxf><font><sz val="10"/></font></dxf>
              <dxf><font><color rgb="FF00FF00"/></font></dxf>
              <dxf><font><color rgb="FF0000FF"/></font></dxf>
            </dxfs>
            <tableStyles count="1"
                         defaultPivotStyle="PivotStyleLight1">
              <tableStyle name="Test Style"
                          pivot="1"
                          count="7">
                <tableStyleElement type="wholeTable" dxfId="0"/>
                <tableStyleElement type="totalRow" dxfId="1"/>
                <tableStyleElement type="firstRowStripe" size="2" dxfId="0"/>
                <tableStyleElement type="secondRowStripe" size="3" dxfId="1"/>
                <tableStyleElement type="firstColumnStripe" size="4" dxfId="2"/>
                <tableStyleElement type="secondColumnStripe" size="5" dxfId="3"/>
                <tableStyleElement type="firstSubtotalColumn" dxfId="2"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            Assert.AreEqual("PivotStyleLight1", styles.DefaultPivotStyle);

            // Style names are case insensitive
            var pivotStyle = styles.PivotStyles["test style"];
            Assert.AreEqual("Test Style", pivotStyle.Name);

            Assert.That(pivotStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLPivotStyleRegionValues, XLDifferentialFormat>
            {
                { XLPivotStyleRegionValues.WholeTable, styles.DifferentialFormats[0] },
                { XLPivotStyleRegionValues.GrandTotalRow, styles.DifferentialFormats[1] },
                { XLPivotStyleRegionValues.FirstRowStripe, styles.DifferentialFormats[0] },
                { XLPivotStyleRegionValues.SecondRowStripe, styles.DifferentialFormats[1] },
                { XLPivotStyleRegionValues.FirstColumnStripe, styles.DifferentialFormats[2] },
                { XLPivotStyleRegionValues.SecondColumnStripe, styles.DifferentialFormats[3] },
                { XLPivotStyleRegionValues.SubtotalColumn1, styles.DifferentialFormats[2] },
            }));

            Assert.AreEqual(2, pivotStyle.RowStripe1BandSize);
            Assert.AreEqual(3, pivotStyle.RowStripe2BandSize);
            Assert.AreEqual(4, pivotStyle.ColumnStripe1BandSize);
            Assert.AreEqual(5, pivotStyle.ColumnStripe2BandSize);
        });
    }

    [Test]
    public void Can_read_pivot_style_with_repeated_elements()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
              <dxf><font><color rgb="FF00FF00"/></font></dxf>
            </dxfs>
            <tableStyles>
              <tableStyle name="Test Style" pivot="1">
                <tableStyleElement type="pageFieldLabels" dxfId="0"/>
                <tableStyleElement type="pageFieldLabels" dxfId="1"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            // When a region is specified multiple times, take the last one
            var pivotStyle = styles.PivotStyles["Test Style"];
            Assert.That(pivotStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLPivotStyleRegionValues, XLDifferentialFormat>
            {
                { XLPivotStyleRegionValues.PageFieldLabels, styles.DifferentialFormats[1] }
            }));
        });
    }

    [Test]
    public void Ignores_pivot_style_elements_that_are_only_for_table_style()
    {
        var xml =
            """
            <dxfs>
              <dxf><font><sz val="5"/></font></dxf>
            </dxfs>
            <tableStyles>
              <tableStyle name="Test Style">
                <tableStyleElement type="lastHeaderCell" dxfId="0"/>
                <tableStyleElement type="lastColumn" dxfId="0"/>
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            var tableStyle = styles.PivotStyles["Test Style"];
            Assert.That(tableStyle.RegionFormats, Is.EquivalentTo(new Dictionary<XLPivotStyleRegionValues, XLDifferentialFormat>
            {
                { XLPivotStyleRegionValues.GrandTotalColumn, styles.DifferentialFormats[0] }
            }));
        });
    }

    [Test]
    public void Ignores_pivot_style_elements_without_differential_format()
    {
        var xml =
            """
            <tableStyles>
              <tableStyle name="Test Style">
                <tableStyleElement type="firstColumn" />
              </tableStyle>
            </tableStyles>
            """;
        AssertTableStyles(xml, styles =>
        {
            var tableStyle = styles.PivotStyles["Test Style"];
            Assert.That(tableStyle.RegionFormats, Is.Empty);
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

    private static void AssertFonts(string fontsXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     <fonts>
                       {fontsXml}
                     </fonts>
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

    private static void AssertDxf(string dxfXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     <dxfs>
                       {dxfXml}
                     </dxfs>
                   </styleSheet>
                   """;
        AssertFormat(assert, xml);
    }

    private void AssertTableStyles(string tableStyleXml, Action<XLWorkbookStyles> assert)
    {
        var xml = $"""
                   <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                     {tableStyleXml}
                   </styleSheet>
                   """;
        AssertFormat(assert, xml);
    }

    private static void AssertFormat(Action<XLWorkbookStyles> assert, string xml)
    {
        using var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(xml));
        using var xmlTreeReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, true);
        var styles = new XLWorkbookStyles();
        var reader = new StylesReader(xmlTreeReader, styles);
        reader.Load();
        assert(styles);
    }
}
