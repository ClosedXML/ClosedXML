using static ClosedXML.Excel.IO.OpenXmlConst;
namespace ClosedXML.Excel.IO;

/// <summary>
/// Reader of style part.
/// </summary>
public class StyleSheetReader // TODO: Make internal, public so I can execute form sandbox
{
    public void Load(XmlTreeReader xml)
    {
        xml.Open("styleSheet", Main2006SsNs);

        if (xml.TryOpen("numFmts", Main2006SsNs))
            ParseNumFmts(xml);

        if (xml.TryOpen("fonts", Main2006SsNs))
            ParseFonts(xml);

        if (xml.TryOpen("fills", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("borders", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("cellStyleXfs", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("cellXfs", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("cellStyles", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen(@"dxfs", Main2006SsNs))
            ParseDxfs(xml);

        if (xml.TryOpen("tableStyles", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("colors", Main2006SsNs))
            xml.Skip();

        if (xml.TryOpen("extLst", Main2006SsNs))
            xml.Skip();

        xml.Close("styleSheet", Main2006SsNs);
    }

    private void ParseNumFmts(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("numFmt", Main2006SsNs))
            ParseNumFmt(xml);

        xml.Close("numFmts", Main2006SsNs);
    }

    private void ParseFonts(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("font", Main2006SsNs))
            ParseFont(xml);

        xml.Close("fonts", Main2006SsNs);
    }

    private void ParseDxfs(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("dxf", Main2006SsNs))
        {
            ParseDxf(xml);
        }

        // Element
        // Here we are, the end element should be 
        xml.Close("dxfs", Main2006SsNs);
    }

    private void ParseDxf(XmlTreeReader xml)
    {
        // I am at the opening element of dfx
        if (xml.TryOpen("font", Main2006SsNs))
            ParseFont(xml);

        if (xml.TryOpen("numFmt", Main2006SsNs))
            ParseNumFmt(xml);

        if (xml.TryOpen("fill", Main2006SsNs))
            ParseFill(xml);

        if (xml.TryOpen("alignment", Main2006SsNs))
            ParseCellAlignment(xml);

        if (xml.TryOpen("border", Main2006SsNs))
            ParseBorder(xml);

        if (xml.TryOpen("protection", Main2006SsNs))
            ParseCellProtection(xml);

        if (xml.TryOpen("extLst", Main2006SsNs))
            xml.Skip();

        xml.Close("dxf", Main2006SsNs);
    }

    private void ParseFont(XmlTreeReader xml)
    {
        // All font elements are optional, so mark them as null by default
        bool? bold = null;
        bool? italic = null;
        bool? strikethrough = null;
        bool? condense = null;
        bool? extend = null;
        bool? outline = null;
        bool? shadow = null;
        XLFontUnderlineValues? underline = null;
        XLFontVerticalTextAlignmentValues? vertAlign = null;
        double? size = null;
        XLColorKey? fontColor = null;
        XLFontFamilyNumberingValues? family = null;
        XLFontCharSet? charSet = null;
        string? name = null;
        XLFontScheme? scheme = null;

        // [MS-OI29500] Excel requires child elements to be in following order: b, i, strike,
        //              condense, extend, outline, shadow, u, vertAlign, sz, color, name, family,
        //              charset, scheme.
        // Official schema doesn't though, it even allows repetition of same element.
        while (!xml.TryClose("font", Main2006SsNs))
        {
            if (xml.TryReadBoolElement("b", out var boldValue))
            {
                bold = boldValue;
            }
            else if (xml.TryReadBoolElement("i", out var italicValue))
            {
                italic = italicValue;
            }
            else if (xml.TryReadBoolElement("strike", out var strikethroughValue))
            {
                strikethrough = strikethroughValue;
            }
            else if (xml.TryReadBoolElement("condense", out var condenseValue))
            {
                condense = condenseValue;
            }
            else if (xml.TryReadBoolElement("extend", out var extendValue))
            {
                extend = extendValue;
            }
            else if (xml.TryReadBoolElement("outline", out var outlineValue))
            {
                outline = outlineValue;
            }
            else if (xml.TryReadBoolElement("shadow", out var shadowValue))
            {
                shadow = shadowValue;
            }
            else if (xml.TryOpen("u", Main2006SsNs))
            {
                underline = xml.GetOptionalEnum("val", XLFontUnderlineValues.Single);
                xml.Close("u", Main2006SsNs);
            }
            else if (xml.TryOpen("vertAlign", Main2006SsNs))
            {
                vertAlign = xml.GetEnum<XLFontVerticalTextAlignmentValues>("val");
                xml.Close("vertAlign", Main2006SsNs);
            }
            else if (xml.TryOpen("sz", Main2006SsNs))
            {
                size = xml.GetDouble("val");
                xml.Close("sz", Main2006SsNs);
            }
            else if (xml.TryParseColor("color", Main2006SsNs, out var color))
            {
                fontColor = color.Key;
            }
            else if (xml.TryOpen("name", Main2006SsNs))
            {
                name = xml.GetAsXString("val");
                xml.Close("name", Main2006SsNs);
            }
            else if (xml.TryOpen("family", Main2006SsNs))
            {
                family = (XLFontFamilyNumberingValues)xml.GetUint("val");
                xml.Close("family", Main2006SsNs);
            }
            else if (xml.TryOpen("charset", Main2006SsNs))
            {
                charSet = (XLFontCharSet)xml.GetInt("val");
                xml.Close("charset", Main2006SsNs);
            }
            else if (xml.TryOpen("scheme", Main2006SsNs))
            {
                scheme = xml.GetEnum<XLFontScheme>("val");
                xml.Close("scheme", Main2006SsNs);
            }
            else
            {
                throw PartStructureException.ExpectedElementNotFound(xml.ElementName);
            }
        }
    }

    private void ParseNumFmt(XmlTreeReader xml)
    {
        xml.Skip();
        //        xml.Close("numFmt", Main2006SsNs);
        //        throw new NotImplementedException();
    }
    private void ParseFill(XmlTreeReader xml)
    {
        xml.Skip();
        //        xml.Close("fill", Main2006SsNs);
        //        throw new NotImplementedException();
    }
    private void ParseCellAlignment(XmlTreeReader xml)
    {
        xml.Skip();
        //        xml.Close("alignment", Main2006SsNs);
        //        throw new NotImplementedException();
    }

    private void ParseBorder(XmlTreeReader xml)
    {
        xml.Skip();
        //        xml.Close("border", Main2006SsNs);
        //        throw new NotImplementedException();
    }

    private void ParseCellProtection(XmlTreeReader xml)
    {
        xml.Skip();
        //        xml.Close("protection", Main2006SsNs);
        //        throw new NotImplementedException();
    }
}
