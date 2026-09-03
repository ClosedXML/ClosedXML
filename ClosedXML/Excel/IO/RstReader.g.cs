#nullable enable

using System.Collections.Generic;
using ClosedXML.Excel.Formatting;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.IO;
using PhoneticProperties = ClosedXML.Excel.XLImmutableRichText.PhoneticProperties;
using PhoneticRunDto = (string Text, int StartIndex, int EndIndex);

namespace ClosedXML.Excel.IO;

internal partial class RstReader
{
    private Xpr<PhoneticRunDto> ParsePhoneticRun(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<PhoneticRunDto>();
        }

        var sb = _reader.GetUInt("sb");
        var eb = _reader.GetUInt("eb");

        var t = _reader.ParseXString("t", _ns).Value;
        _reader.Close(elementName, ns);

        return Xpr.From(OnPhoneticRunParsed(t, sb, eb));
    }

    private Xpr<(string Text, XLDifferentialFontValue Font)> ParseRElt(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<(string Text, XLDifferentialFontValue Font)>();
        }

        var rPrResult = ParseRPrElt("rPr", _ns);
        var rPr = rPrResult.IsSuccess ? rPrResult.Value : default(XLDifferentialFontValue?);
        var t = _reader.ParseXString("t", _ns).Value;
        _reader.Close(elementName, ns);

        return Xpr.From(OnREltParsed(rPr, t));
    }

    private Xpr<XLDifferentialFontValue> ParseRPrElt(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLDifferentialFontValue>();
        }

        // Choice with cardinality 1-n
        var choiceList = new List<Unit>();
        var choiceCount = 0;
        while (true)
        {
            Unit choice;
            if (ParseFontName("rFont", _ns) is { IsSuccess: true } rFont)
            {
                choice = OnRPrEltRFontParsed(rFont.Value);
            }
            else if (ParseIntProperty("charset", _ns) is { IsSuccess: true } charset)
            {
                choice = OnRPrEltCharsetParsed(charset.Value);
            }
            else if (ParseIntProperty("family", _ns) is { IsSuccess: true } family)
            {
                choice = OnRPrEltFamilyParsed(family.Value);
            }
            else if (ParseBooleanProperty("b", _ns) is { IsSuccess: true } b)
            {
                choice = OnRPrEltBParsed(b.Value);
            }
            else if (ParseBooleanProperty("i", _ns) is { IsSuccess: true } i)
            {
                choice = OnRPrEltIParsed(i.Value);
            }
            else if (ParseBooleanProperty("strike", _ns) is { IsSuccess: true } strike)
            {
                choice = OnRPrEltStrikeParsed(strike.Value);
            }
            else if (ParseBooleanProperty("outline", _ns) is { IsSuccess: true } outline)
            {
                choice = OnRPrEltOutlineParsed(outline.Value);
            }
            else if (ParseBooleanProperty("shadow", _ns) is { IsSuccess: true } shadow)
            {
                choice = OnRPrEltShadowParsed(shadow.Value);
            }
            else if (ParseBooleanProperty("condense", _ns) is { IsSuccess: true } condense)
            {
                choice = OnRPrEltCondenseParsed(condense.Value);
            }
            else if (ParseBooleanProperty("extend", _ns) is { IsSuccess: true } extend)
            {
                choice = OnRPrEltExtendParsed(extend.Value);
            }
            else if (ParseColor("color", _ns) is { IsSuccess: true } color)
            {
                choice = OnRPrEltColorParsed(color.Value);
            }
            else if (ParseFontSize("sz", _ns) is { IsSuccess: true } sz)
            {
                choice = OnRPrEltSzParsed(sz.Value);
            }
            else if (ParseUnderlineProperty("u", _ns) is { IsSuccess: true } u)
            {
                choice = OnRPrEltUParsed(u.Value);
            }
            else if (ParseVerticalAlignFontProperty("vertAlign", _ns) is { IsSuccess: true } vertAlign)
            {
                choice = OnRPrEltVertAlignParsed(vertAlign.Value);
            }
            else if (ParseFontScheme("scheme", _ns) is { IsSuccess: true } scheme)
            {
                choice = OnRPrEltSchemeParsed(scheme.Value);
            }
            else
            {
                break;
            }
            choiceList.Add(choice);
            choiceCount++;
        }
        if(choiceCount == 0)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        return Xpr.From(OnRPrEltParsed(choiceList));
    }

    private Xpr<OneOf<string, XLImmutableRichText>> ParseRst(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<OneOf<string, XLImmutableRichText>>();
        }

        var tResult = _reader.ParseXString("t", _ns);
        var t = tResult.IsSuccess ? tResult.Value : default(string?);
        var r = new List<(string Text, XLDifferentialFontValue Font)>();
        while (ParseRElt("r", _ns) is { IsSuccess: true} rItem)
        {
            r.Add(rItem.Value);
        }
        var rPh = new List<PhoneticRunDto>();
        while (ParsePhoneticRun("rPh", _ns) is { IsSuccess: true} rPhItem)
        {
            rPh.Add(rPhItem.Value);
        }
        var phoneticPrResult = ParsePhoneticPr("phoneticPr", _ns);
        var phoneticPr = phoneticPrResult.IsSuccess ? phoneticPrResult.Value : default(PhoneticProperties?);
        _reader.Close(elementName, ns);

        return Xpr.From(OnRstParsed(t, r, rPh, phoneticPr));
    }

    private Xpr<PhoneticProperties> ParsePhoneticPr(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<PhoneticProperties>();
        }

        var fontId = _reader.GetUInt("fontId");
        var type = _reader.GetOptionalEnum<XLPhoneticType>("type") ?? XLPhoneticType.FullWidthKatakana;
        var alignment = _reader.GetOptionalEnum<XLPhoneticAlignment>("alignment") ?? XLPhoneticAlignment.Left;

        _reader.Close(elementName, ns);

        return Xpr.From(OnPhoneticPrParsed(fontId, type, alignment));
    }

    private Xpr<bool> ParseBooleanProperty(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<bool>();
        }

        var val = _reader.GetOptionalBool("val") ?? true;

        _reader.Close(elementName, ns);

        return Xpr.From(OnBooleanPropertyParsed(val));
    }

    private Xpr<XLFontSize> ParseFontSize(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFontSize>();
        }

        var val = _reader.GetDouble("val");

        _reader.Close(elementName, ns);

        return Xpr.From(OnFontSizeParsed(val));
    }

    private Xpr<int> ParseIntProperty(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<int>();
        }

        var val = _reader.GetInt("val");

        _reader.Close(elementName, ns);

        return Xpr.From(OnIntPropertyParsed(val));
    }

    private Xpr<XLFontName> ParseFontName(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFontName>();
        }

        var val = _reader.GetXString("val");

        _reader.Close(elementName, ns);

        return Xpr.From(OnFontNameParsed(val));
    }

    private Xpr<XLFontVerticalTextAlignmentValues> ParseVerticalAlignFontProperty(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFontVerticalTextAlignmentValues>();
        }

        var val = _reader.GetEnum<XLFontVerticalTextAlignmentValues>("val");

        _reader.Close(elementName, ns);

        return Xpr.From(OnVerticalAlignFontPropertyParsed(val));
    }

    private Xpr<XLFontScheme> ParseFontScheme(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFontScheme>();
        }

        var val = _reader.GetEnum<XLFontScheme>("val");

        _reader.Close(elementName, ns);

        return Xpr.From(OnFontSchemeParsed(val));
    }

    private Xpr<XLFontUnderlineValues> ParseUnderlineProperty(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLFontUnderlineValues>();
        }

        var val = _reader.GetOptionalEnum<XLFontUnderlineValues>("val") ?? XLFontUnderlineValues.Single;

        _reader.Close(elementName, ns);

        return Xpr.From(OnUnderlinePropertyParsed(val));
    }
}
