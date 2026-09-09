using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.Formatting;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using static ClosedXML.Excel.XLImmutableRichText;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class RstReaderTests
{
    [Test]
    public void Empty_si_is_an_empty_string()
    {
        AssertPlainText("<si/>", string.Empty);
    }

    [Test]
    public void Can_read_plain_text()
    {
        AssertPlainText("""
            <si>
              <t>Hello</t>
            </si>
            """,
            "Hello");
    }

    [Test]
    public void Empty_t_element_is_an_empty_string()
    {
        AssertPlainText("""
            <si>
              <t/>
            </si>
            """,
            string.Empty);
    }

    [Test]
    public void Decodes_xstring_in_plain_text()
    {
        AssertPlainText("""
            <si>
              <t>A_x0009_B</t>
            </si>
            """,
            "A\tB");
    }

    [Test]
    public void Parse_rst_method_fails_if_element_is_not_found()
    {
        ParseItems(
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
            (rstReader, xmlReader) =>
            {
                xmlReader.Open("sst", OpenXmlConst.Main2006SsNs);
                var result = rstReader.ParseCtRst("si", OpenXmlConst.Main2006SsNs);
                Assert.True(result.IsFail);
            });
    }

    [Test]
    public void Can_read_unstyled_rich_run()
    {
        var richText = ParseRichText("""
            <si>
              <r>
                <t>Hello</t>
              </r>
            </si>
            """);

        Assert.AreEqual("Hello", richText.Text);
        Assert.AreEqual(richText.Runs, new[] { new RichTextRun(XLDifferentialFontValue.Empty, 0, 5) });
        Assert.AreEqual(0, richText.PhoneticRuns.Count);
        Assert.IsNull(richText.PhoneticsProperties);
    }

    [Test]
    public void Concatenates_multiple_rich_runs()
    {
        var richText = ParseRichText("""
            <si>
              <r>
                <rPr>
                  <b/>
                </rPr>
                <t>Hel</t>
              </r>
              <r>
                <t>lo</t>
              </r>
            </si>
            """);

        Assert.AreEqual("Hello", richText.Text);
        Assert.AreEqual(richText.Runs, new[]
        {
            new RichTextRun(XLDifferentialFontValue.Empty with { Bold = true }, 0, 3),
            new RichTextRun(XLDifferentialFontValue.Empty, 3, 2),
        });
    }

    [Test]
    public void Empty_rich_run_is_still_rich_text()
    {
        var richText = ParseRichText("""
            <si>
              <r>
                <t/>
              </r>
            </si>
            """);

        Assert.AreEqual(string.Empty, richText.Text);
        Assert.AreEqual(1, richText.Runs.Count);
        Assert.AreEqual(0, richText.Runs[0].Length);
    }

    [Test]
    public void Plain_text_element_is_prepended_as_initial_rich_run()
    {
        var richText = ParseRichText("""
            <si>
              <t>Head</t>
              <r>
                <t>Tail</t>
              </r>
            </si>
            """);

        Assert.AreEqual("HeadTail", richText.Text);
        Assert.AreEqual(2, richText.Runs.Count);
        Assert.AreEqual("Head", richText.GetRunText(richText.Runs[0]));
        Assert.AreEqual("Tail", richText.GetRunText(richText.Runs[1]));
        Assert.AreEqual(XLDifferentialFontValue.Empty, richText.Runs[1].Dxf);
    }

    [Test]
    public void Empty_plain_text_is_not_prepended_to_rich_runs()
    {
        var richText = ParseRichText("""
            <si>
              <t/>
              <r>
                <t>Only</t>
              </r>
            </si>
            """);

        Assert.AreEqual("Only", richText.Text);
        Assert.AreEqual(1, richText.Runs.Count);
    }

    [Test]
    public void Run_font_state_is_reset_between_runs()
    {
        var richText = ParseRichText("""
            <si>
              <r>
                <rPr>
                  <b/>
                </rPr>
                <t>A</t>
              </r>
              <r>
                <t>B</t>
              </r>
            </si>
            """);

        Assert.IsTrue(richText.Runs[0].Dxf.Bold);
        Assert.IsNull(richText.Runs[1].Dxf.Bold);
    }

    [Test]
    public void Run_font_state_is_reset_between_rst_elements()
    {
        ParseItems($"""
            <sst xmlns="{OpenXmlConst.Main2006SsNs}">
              <si>
                <r>
                  <rPr>
                    <b/>
                  </rPr>
                  <t>A</t>
                </r>
              </si>
              <si>
                <r>
                  <t>B</t>
                </r>
              </si>
            </sst>
            """,
            (rstReader, xmlReader) =>
            {
                xmlReader.Open("sst", OpenXmlConst.Main2006SsNs);
                var first = rstReader.ParseCtRst("si", OpenXmlConst.Main2006SsNs).Value;
                var second = rstReader.ParseCtRst("si", OpenXmlConst.Main2006SsNs).Value;

                Assert.IsTrue(GetRichText(first).Runs[0].Dxf.Bold);
                Assert.IsNull(GetRichText(second).Runs[0].Dxf.Bold);
            });
    }

    [Test]
    public void Can_read_run_properties()
    {
        // Excel requires child elements in this sequence. Boolean attributes mix omitted val
        // (defaults to true) and an explicit false.
        var richText = ParseRichText("""
            <si>
              <r>
                <rPr>
                  <b/>
                  <i val="0"/>
                  <strike/>
                  <condense val="0"/>
                  <extend/>
                  <outline/>
                  <shadow val="0"/>
                  <u val="double"/>
                  <vertAlign val="superscript"/>
                  <sz val="8.5"/>
                  <color rgb="FF802010"/>
                  <rFont val="Calibri"/>
                  <family val="2"/>
                  <charset val="128"/>
                  <scheme val="none"/>
                </rPr>
                <t>Text</t>
              </r>
            </si>
            """);

        var dxf = richText.Runs[0].Dxf;
        Assert.AreEqual("Calibri", dxf.Name?.Text);
        Assert.AreEqual(XLFontCharSet.ShiftJIS, dxf.Charset);
        Assert.AreEqual(XLFontFamilyNumberingValues.Swiss, dxf.Family);
        Assert.IsTrue(dxf.Bold);
        Assert.IsFalse(dxf.Italic);
        Assert.IsTrue(dxf.Strikethrough);
        Assert.IsTrue(dxf.Outline);
        Assert.IsFalse(dxf.Shadow);
        Assert.IsFalse(dxf.Condense);
        Assert.IsTrue(dxf.Extend);
        Assert.AreEqual(XLColor.FromRgb(0x802010), dxf.Color);
        Assert.AreEqual(8.5, dxf.Size);
        Assert.AreEqual(XLFontUnderlineValues.Double, dxf.Underline);
        Assert.AreEqual(XLFontVerticalTextAlignmentValues.Superscript, dxf.VerticalAlignment);
        Assert.AreEqual(XLFontScheme.None, dxf.Scheme);
    }

    [Test]
    public void Underline_without_value_defaults_to_single()
    {
        var richText = ParseRichText("""
            <si>
              <r>
                <rPr>
                  <u/>
                </rPr>
                <t>Text</t>
              </r>
            </si>
            """);

        Assert.AreEqual(XLFontUnderlineValues.Single, richText.Runs[0].Dxf.Underline);
    }

    [TestCase(6)]
    [TestCase(14)]
    public void Undefined_font_family_value_is_interpreted_as_unknown_font_family(int fontFamily)
    {
        // OI-29500: Excel restricts the value of this attribute to be at least 0 and at most 5.
        var richText = ParseRichText($"""
            <si>
              <r>
                <rPr>
                  <family val="{fontFamily}"/>
                </rPr>
                <t>Text</t>
              </r>
            </si>
            """);

        Assert.AreEqual(XLFontFamilyNumberingValues.NotApplicable, richText.Runs[0].Dxf.Family);
    }

    [TestCase(-1)]
    [TestCase(15)]
    public void Rejects_font_family_outside_spec_range(int invalidFontFamily)
    {
        Assert.That(
            () => ParseRichText($"""
                <si>
                  <r>
                    <rPr>
                      <family val="{invalidFontFamily}"/>
                    </rPr>
                    <t>Text</t>
                  </r>
                </si>
                """),
            Throws.Exception.TypeOf<PartStructureException>().And
                .Message.StartsWith(PartStructureException.InvalidAttributeValue(invalidFontFamily.ToString()).Message));
    }

    [TestCase(-1)]
    [TestCase(256)]
    public void Rejects_charset_outside_byte_range(int invalidCharset)
    {
        Assert.That(
            () => ParseRichText($"""
                <si>
                  <r>
                    <rPr>
                      <charset val="{invalidCharset}"/>
                    </rPr>
                    <t>Text</t>
                  </r>
                </si>
                """),
            Throws.Exception.TypeOf<PartStructureException>().And
                .Message.StartsWith(PartStructureException.InvalidAttributeValue(invalidCharset.ToString()).Message));
    }

    [Test]
    public void Run_properties_must_have_at_least_one_font_component()
    {
        Assert.That(
            () => ParseRichText("""
                <si>
                  <r>
                    <rPr/>
                    <t>A</t>
                  </r>
                </si>
                """),
            Throws.Exception.TypeOf<PartStructureException>().And
                .Message.StartsWith(PartStructureException.IncorrectElementsCount().Message));
    }

    [Test]
    public void Can_read_phonetic_runs_and_properties()
    {
        var phoneticFont = XLFontFormatValue.Default with { Name = @"Meiryo", Size = XLFontSize.FromPoints(6) };
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>東京</t>
              </r>
              <rPh sb="0" eb="2">
                <t>とうきょう</t>
              </rPh>
              <phoneticPr fontId="0" type="Hiragana" alignment="distributed"/>
            </si>
            """,
            styles => styles.AddFontFormat(phoneticFont));

        Assert.AreEqual("東京", richText.Text);
        Assert.That(richText.PhoneticRuns, Is.EqualTo([new PhoneticRun("とうきょう", 0, 2)]));
        Assert.That(richText.PhoneticsProperties, Is.EqualTo(new PhoneticProperties(phoneticFont, XLPhoneticType.Hiragana, XLPhoneticAlignment.Distributed)));
    }

    [Test]
    public void Phonetic_runs_are_sorted_by_start_index()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>abcd</t>
              </r>
              <rPh sb="2" eb="4">
                <t>CD</t>
              </rPh>
              <rPh sb="0" eb="2">
                <t>AB</t>
              </rPh>
            </si>
            """);

        Assert.That(richText.PhoneticRuns, Is.EqualTo(
            [
                new PhoneticRun("AB", 0, 2),
                new PhoneticRun("CD", 2, 4)
            ]));
    }

    [Test]
    public void Overlapping_phonetic_runs_are_dropped()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>abcd</t>
              </r>
              <rPh sb="0" eb="3">
                <t>ABC</t>
              </rPh>
              <rPh sb="2" eb="4">
                <t>CD</t>
              </rPh>
            </si>
            """);

        Assert.That(richText.PhoneticRuns, Is.EqualTo([new PhoneticRun("CD", 2, 4)]));
    }

    [Test]
    public void Adjacent_phonetic_runs_are_kept()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>abcd</t>
              </r>
              <rPh sb="0" eb="2">
                <t>AB</t>
              </rPh>
              <rPh sb="2" eb="4">
                <t>CD</t>
              </rPh>
            </si>
            """);

        Assert.AreEqual(2, richText.PhoneticRuns.Count);
    }

    [Test]
    public void Phonetic_runs_with_out_of_bound_index_are_filtered_out()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>ab</t>
              </r>
              <rPh sb="0" eb="4">
                <t>ABCD</t>
              </rPh>
            </si>
            """);

        Assert.AreEqual(0, richText.PhoneticRuns.Count);
    }

    [Test]
    public void Phonetic_runs_with_start_index_greater_or_equal_to_end_index_are_filtered_out()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>Text</t>
              </r>
              <rPh sb="1" eb="0">
                <t>Omitted</t>
              </rPh>
              <rPh sb="2" eb="2">
                <t>Omitted</t>
              </rPh>
            </si>
            """);

        Assert.AreEqual(0, richText.PhoneticRuns.Count);
    }

    [Test]
    public void Phonetic_runs_without_text_are_filtered_out()
    {
        var richText = ParseRichText(
            """
            <si>
              <r>
                <t>Text</t>
              </r>
              <rPh sb="0" eb="4">
                <t/>
              </rPh>
            </si>
            """);

        Assert.AreEqual(0, richText.PhoneticRuns.Count);
    }

    [Test]
    public void Unknown_phonetic_font_id_throws()
    {
        Assert.That(
            () => ParseRichText("""
                <si>
                  <phoneticPr fontId="30"/>
                </si>
                """),
            Throws.TypeOf<KeyNotFoundException>());
    }

    private static void AssertPlainText(string siXml, string expected)
    {
        var parsed = ParseSi(siXml);
        Assert.True(parsed.TryPickT0(out var text, out _));
        Assert.AreEqual(expected, text);
    }

    private static XLImmutableRichText ParseRichText(string siXml, Action<XLWorkbookStyles> configureStyles = null)
    {
        return GetRichText(ParseSi(siXml, configureStyles));
    }

    private static XLImmutableRichText GetRichText(OneOf<string, XLImmutableRichText> parsed)
    {
        Assert.False(parsed.TryPickT0(out _, out var richText));
        Assert.NotNull(richText);
        return richText;
    }

    private static OneOf<string, XLImmutableRichText> ParseSi(string xml, Action<XLWorkbookStyles> configureStyles = null)
    {
        OneOf<string, XLImmutableRichText>? parsed = null;
        ParseItems(
            $"""
             <sst xmlns="{OpenXmlConst.Main2006SsNs}">
               {xml}
             </sst>
             """,
            (rstReader, xmlReader) =>
            {
                xmlReader.Open("sst", OpenXmlConst.Main2006SsNs);
                parsed = rstReader.ParseCtRst("si", OpenXmlConst.Main2006SsNs).Value;
            },
            configureStyles);

        Assert.NotNull(parsed);
        return parsed.Value;
    }

    private static void ParseItems(
        string xml,
        Action<RstReader, XmlTreeReader> act,
        Action<XLWorkbookStyles> configureStyles = null)
    {
        using var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(xml));
        using var xmlReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, true);
        var styles = new XLWorkbookStyles();
        configureStyles?.Invoke(styles);
        var rstReader = new RstReader(xmlReader, styles);
        act(rstReader, xmlReader);
    }
}
