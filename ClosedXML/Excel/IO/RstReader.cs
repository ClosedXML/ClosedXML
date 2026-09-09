using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.Formatting;
using ClosedXML.IO;
using PhoneticProperties = ClosedXML.Excel.XLImmutableRichText.PhoneticProperties;
using PhoneticRun = ClosedXML.Excel.XLImmutableRichText.PhoneticRun;
using RichTextRun = ClosedXML.Excel.XLImmutableRichText.RichTextRun;
using PhoneticRunDto = (string Text, int StartIndex, int EndIndex);

namespace ClosedXML.Excel.IO;

/// <summary>
/// A partial stateful reader for reading a string item (it's either text or rich text). The rich
/// text format is not stored in the <see cref="XLWorkbookStyles"/> and the reader doesn't add it
/// there.
/// </summary>
internal partial class RstReader
{
    private static readonly Comparer<PhoneticRunDto> ComparePhoneticRunsByStartIndex = Comparer<PhoneticRunDto>.Create((x, y) => x.StartIndex.CompareTo(y.StartIndex));

    private readonly string _ns = OpenXmlConst.Main2006SsNs;

    private readonly XmlTreeReader _reader;

    private readonly XLWorkbookStyles _styles;

    // The values read from the CT_RPrElt element. The values are reset before each run.
    private XLFontName? _runName;
    private XLFontCharSet? _runCharset;
    private XLFontFamilyNumberingValues? _runFamily;
    private bool? _runBold;
    private bool? _runItalic;
    private bool? _runStrikethrough;
    private bool? _runOutline;
    private bool? _runShadow;
    private bool? _runCondense;
    private bool? _runExtend;
    private XLColor? _runColor;
    private XLFontSize? _runSize;
    private XLFontUnderlineValues? _runUnderline;
    private XLFontVerticalTextAlignmentValues? _runVerticalAlignment;
    private XLFontScheme? _runScheme;

    internal RstReader(XmlTreeReader reader, XLWorkbookStyles styles)
    {
        _reader = reader;
        _styles = styles;
    }

    /// <summary>
    /// Parse a <c>CT_Rst</c> element.
    /// </summary>
    internal Xpr<OneOf<string, XLImmutableRichText>> ParseCtRst(string elementName, string ns)
    {
        ResetFontState();
        return ParseRst(elementName, ns);
    }

    /// <summary>
    /// Reset information about properties of current/next rich run.
    /// </summary>
    private void ResetFontState()
    {
        _runName = null;
        _runCharset = null;
        _runFamily = null;
        _runBold = null;
        _runItalic = null;
        _runStrikethrough = null;
        _runOutline = null;
        _runShadow = null;
        _runCondense = null;
        _runExtend = null;
        _runColor = null;
        _runSize = null;
        _runUnderline = null;
        _runVerticalAlignment = null;
        _runScheme = null;
    }

    /// <summary>
    /// Get the state of properties for current/next rich run.
    /// </summary>
    private XLDifferentialFontValue GetFontState()
    {
        return new XLDifferentialFontValue
        {
            Name = _runName,
            Charset = _runCharset,
            Family = _runFamily,
            Bold = _runBold,
            Italic = _runItalic,
            Strikethrough = _runStrikethrough,
            Outline = _runOutline,
            Shadow = _runShadow,
            Condense = _runCondense,
            Extend = _runExtend,
            Color = _runColor,
            Size = _runSize,
            Underline = _runUnderline,
            VerticalAlignment = _runVerticalAlignment,
            Scheme = _runScheme
        };
    }

    private bool OnBooleanPropertyParsed(bool value)
    {
        return value;
    }

    private int OnIntPropertyParsed(int value)
    {
        return value;
    }

    private XLFontSize OnFontSizeParsed(double sizePt)
    {
        return XLFontSize.FromPoints(sizePt);
    }

    private XLFontName OnFontNameParsed(string fontName)
    {
        return fontName;
    }

    private XLFontVerticalTextAlignmentValues OnVerticalAlignFontPropertyParsed(XLFontVerticalTextAlignmentValues verticalAlignment)
    {
        return verticalAlignment;
    }

    private XLFontScheme OnFontSchemeParsed(XLFontScheme scheme)
    {
        return scheme;
    }

    private XLFontUnderlineValues OnUnderlinePropertyParsed(XLFontUnderlineValues underline)
    {
        return underline;
    }


    #region Callbacks for the OnRPrElt that store the rich run properties

    private Unit OnRPrEltRFontParsed(XLFontName name)
    {
        _runName = name;
        return Unit.Value;
    }

    private Unit OnRPrEltCharsetParsed(int charsetValue)
    {
        if (charsetValue is < 0 or > 255)
            throw PartStructureException.InvalidAttributeValue(charsetValue.ToString());

        _runCharset = (XLFontCharSet)charsetValue;
        return Unit.Value;
    }

    private Unit OnRPrEltFamilyParsed(int familyValue)
    {
        // Unlike family in the CT_Font, the family in the CT_RPrElt is stored as an int.
        // Bug in the spec. Spec says that it has values 0-14 and doesn't specify meaning
        // for the numerical values. It's supposed to refer to the same enum ST_FontFamily
        // as in WordML. The OI-29500 fixes this problem:
        // "Excel restricts the value of this attribute to be at least 0 and at most 5."
        _runFamily = familyValue switch
        {
            >= 0 and <= 5 => (XLFontFamilyNumberingValues)familyValue,
            > 5 and <= 14 => XLFontFamilyNumberingValues.NotApplicable,
            _ => throw PartStructureException.InvalidAttributeValue(familyValue.ToString()),
        };
        return Unit.Value;
    }

    private Unit OnRPrEltBParsed(bool bold)
    {
        _runBold = bold;
        return Unit.Value;
    }

    private Unit OnRPrEltIParsed(bool italic)
    {
        _runItalic = italic;
        return Unit.Value;
    }

    private Unit OnRPrEltStrikeParsed(bool strikethrough)
    {
        _runStrikethrough = strikethrough;
        return Unit.Value;
    }

    private Unit OnRPrEltOutlineParsed(bool outline)
    {
        _runOutline = outline;
        return Unit.Value;
    }

    private Unit OnRPrEltShadowParsed(bool shadow)
    {
        _runShadow = shadow;
        return Unit.Value;
    }

    private Unit OnRPrEltCondenseParsed(bool condense)
    {
        _runCondense = condense;
        return Unit.Value;
    }

    private Unit OnRPrEltExtendParsed(bool extend)
    {
        _runExtend = extend;
        return Unit.Value;
    }

    private Xpr<XLColor> ParseColor(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLColor>();
        }

        return Xpr.From(_reader.ParseColor(elementName, ns));
    }

    private Unit OnRPrEltColorParsed(XLColor color)
    {
        _runColor = color;
        return Unit.Value;
    }

    private Unit OnRPrEltSzParsed(XLFontSize size)
    {
        _runSize = size;
        return Unit.Value;
    }

    private Unit OnRPrEltUParsed(XLFontUnderlineValues underline)
    {
        _runUnderline = underline;
        return Unit.Value;
    }

    private Unit OnRPrEltVertAlignParsed(XLFontVerticalTextAlignmentValues verticalAlignment)
    {
        _runVerticalAlignment = verticalAlignment;
        return Unit.Value;
    }

    private Unit OnRPrEltSchemeParsed(XLFontScheme scheme)
    {
        _runScheme = scheme;
        return Unit.Value;
    }

    /// <summary>
    /// Create font format from the values of the current state of the run (the <c>_run*</c> fields).
    /// </summary>
    private XLDifferentialFontValue OnRPrEltParsed(List<Unit> _)
    {
        return GetFontState();
    }

    #endregion

    private (string Text, XLDifferentialFontValue Font) OnREltParsed(XLDifferentialFontValue? runProperties, string runText)
    {
        var runFont = runProperties ?? XLDifferentialFontValue.Empty;
        ResetFontState();
        return (runText, runFont);
    }

    private PhoneticRunDto OnPhoneticRunParsed(string text, uint sb, uint eb)
    {
        // Validate and filter out later, in the method that construct the rich text
        return new PhoneticRunDto(text, checked((int)sb), checked((int)eb));
    }

    private PhoneticProperties OnPhoneticPrParsed(uint fontId, XLPhoneticType type, XLPhoneticAlignment alignment)
    {
        var phoneticFont = _styles.Fonts[checked((int)fontId)];
        return new PhoneticProperties(phoneticFont, type, alignment);
    }

    private OneOf<string, XLImmutableRichText> OnRstParsed(string? fullText, List<(string Text, XLDifferentialFontValue Font)> richRuns, List<PhoneticRunDto> phoneticRuns, PhoneticProperties? phoneticProps)
    {
        if (richRuns.Count == 0 && phoneticProps is null)
            return fullText ?? string.Empty;

        if (!string.IsNullOrEmpty(fullText))
            richRuns.Insert(0, (fullText, XLDifferentialFontValue.Empty));

        return GetRichText(richRuns, phoneticRuns, phoneticProps);
    }

    private static XLImmutableRichText GetRichText(List<(string Text, XLDifferentialFontValue Font)> richRuns, List<PhoneticRunDto> phoneticRunDtos, PhoneticProperties? phoneticProps)
    {
        var richTextLength = richRuns.Sum(x => x.Text.Length);
        var sb = new StringBuilder(richTextLength);
        foreach (var richRun in richRuns)
            sb.Append(richRun.Text);

        var richText = sb.ToString();

        // Phonetic runs must be in ascending order, non-overlapping
        // Any two consecutive phonetic runs should satisfy sb1 < eb1 <= sb2 < eb2 to avoid overlapping
        phoneticRunDtos.Sort(ComparePhoneticRunsByStartIndex);
        var prevStartIndex = richTextLength;
        for (var i = phoneticRunDtos.Count - 1; i >= 0; --i)
        {
            var phoneticRun = phoneticRunDtos[i];

            // Each run must have a text
            if (string.IsNullOrEmpty(phoneticRun.Text))
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Start index must be less than the end index
            if (phoneticRun.EndIndex < phoneticRun.StartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Omit phonetic runs of length 0
            if (phoneticRun.EndIndex == phoneticRun.StartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            // Omit phonetic runs that overlap with the next one
            if (phoneticRun.EndIndex > prevStartIndex)
            {
                phoneticRunDtos.RemoveAt(i);
                continue;
            }

            prevStartIndex = phoneticRun.StartIndex;
        }

        var runs = new RichTextRun[richRuns.Count];
        var startIndex = 0;
        for (var i = 0; i < richRuns.Count; ++i)
        {
            var runLength = richRuns[i].Text.Length;
            runs[i] = new RichTextRun(XLFontFormatValue.Default, richRuns[i].Font, startIndex, runLength);
            startIndex += runLength;
        }

        var phoneticRuns = phoneticRunDtos.Select(x => new PhoneticRun(x.Text, x.StartIndex, x.EndIndex)).ToArray();
        return new XLImmutableRichText(richText, runs, phoneticRuns, phoneticProps);
    }
}
