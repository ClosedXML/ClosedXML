namespace ClosedXML.Excel.IO;

/// <summary>
/// Reader of style part.
/// </summary>
public class StyleSheetReader // TODO: Make internal, public so I can execute form sandbox
{
    private string _mainNs = OpenXmlConst.Main2006SsNs;

    public void Load(XmlTreeReader xml)
    {
        if (!xml.TryOpen("styleSheet", _mainNs))
        {
            // Try OOXML strict namespace
            _mainNs = "http://purl.oclc.org/ooxml/spreadsheetml/main";
            xml.Open("styleSheet", _mainNs);
        }

        if (xml.TryOpen("numFmts", _mainNs))
            ParseNumFmts(xml);

        if (xml.TryOpen("fonts", _mainNs))
            ParseFonts(xml);

        if (xml.TryOpen("fills", _mainNs))
            xml.Skip();

        if (xml.TryOpen("borders", _mainNs))
            xml.Skip();
        
        if (xml.TryOpen("cellStyleXfs", _mainNs))
            xml.Skip();

        if (xml.TryOpen("cellXfs", _mainNs))
            xml.Skip();

        if (xml.TryOpen("cellStyles", _mainNs))
            xml.Skip();

        if (xml.TryOpen(@"dxfs", _mainNs))
            ParseDxfs(xml);

        if (xml.TryOpen("tableStyles", _mainNs))
            xml.Skip();

        if (xml.TryOpen("colors", _mainNs))
            xml.Skip();

        if (xml.TryOpen("extLst", _mainNs))
            xml.Skip();

        xml.Close("styleSheet", _mainNs);
    }

    private void ParseNumFmts(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("numFmt", _mainNs))
            ParseNumFmt(xml);

        xml.Close("numFmts", _mainNs);
    }

    private void ParseFonts(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("font", _mainNs))
            ParseFont(xml);

        xml.Close("fonts", _mainNs);
    }

    private void ParseDxfs(XmlTreeReader xml)
    {
        _ = xml.GetOptionalUint("count");
        while (xml.TryOpen("dxf", _mainNs))
            ParseDxf(xml);

        // Element
        // Here we are, the end element should be 
        xml.Close("dxfs", _mainNs);
    }

    private void ParseDxf(XmlTreeReader xml)
    {
        // I am at the opening element of dfx
        if (xml.TryOpen("font", _mainNs))
            ParseFont(xml);

        if (xml.TryOpen("numFmt", _mainNs))
            ParseNumFmt(xml);

        if (xml.TryOpen("fill", _mainNs))
            ParseFill(xml);

        if (xml.TryOpen("alignment", _mainNs))
            ParseCellAlignment(xml);

        if (xml.TryOpen("border", _mainNs))
            ParseBorder(xml);

        if (xml.TryOpen("protection", _mainNs))
            ParseCellProtection(xml);

        if (xml.TryOpen("extLst", _mainNs))
            xml.Skip();

        xml.Close("dxf", _mainNs);
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
        while (!xml.TryClose("font", _mainNs))
        {
            if (xml.TryReadBoolElement("b", _mainNs, out var boldValue))
            {
                bold = boldValue;
            }
            else if (xml.TryReadBoolElement("i", _mainNs, out var italicValue))
            {
                italic = italicValue;
            }
            else if (xml.TryReadBoolElement("strike", _mainNs, out var strikethroughValue))
            {
                strikethrough = strikethroughValue;
            }
            else if (xml.TryReadBoolElement("condense", _mainNs, out var condenseValue))
            {
                condense = condenseValue;
            }
            else if (xml.TryReadBoolElement("extend", _mainNs, out var extendValue))
            {
                extend = extendValue;
            }
            else if (xml.TryReadBoolElement("outline", _mainNs, out var outlineValue))
            {
                outline = outlineValue;
            }
            else if (xml.TryReadBoolElement("shadow", _mainNs, out var shadowValue))
            {
                shadow = shadowValue;
            }
            else if (xml.TryOpen("u", _mainNs))
            {
                underline = xml.GetOptionalEnum("val", XLFontUnderlineValues.Single);
                xml.Close("u", _mainNs);
            }
            else if (xml.TryOpen("vertAlign", _mainNs))
            {
                vertAlign = xml.GetEnum<XLFontVerticalTextAlignmentValues>("val");
                xml.Close("vertAlign", _mainNs);
            }
            else if (xml.TryOpen("sz", _mainNs))
            {
                size = xml.GetDouble("val");
                xml.Close("sz", _mainNs);
            }
            else if (xml.TryParseColor("color", _mainNs, out var color))
            {
                fontColor = color.Key;
            }
            else if (xml.TryOpen("name", _mainNs))
            {
                name = xml.GetAsXString("val");
                xml.Close("name", _mainNs);
            }
            else if (xml.TryOpen("family", _mainNs))
            {
                family = (XLFontFamilyNumberingValues)xml.GetUint("val");
                xml.Close("family", _mainNs);
            }
            else if (xml.TryOpen("charset", _mainNs))
            {
                charSet = (XLFontCharSet)xml.GetInt("val");
                xml.Close("charset", _mainNs);
            }
            else if (xml.TryOpen("scheme", _mainNs))
            {
                scheme = xml.GetEnum<XLFontScheme>("val");
                xml.Close("scheme", _mainNs);
            }
            else if (xml.TryOpen("scheme", _mainNs))
            {
                scheme = xml.GetEnum<XLFontScheme>("val");
                xml.Close("scheme", _mainNs);
            }
            else
            {
                // TODO: Add option to skip unknown elements. Basically lax parsing. Most XML is well behaved, then... there are screwups. 
                throw PartStructureException.ExpectedElementNotFound(xml.ElementName);
            }
        }
    }

    private void ParseNumFmt(XmlTreeReader xml)
    {
        xml.Skip();
    }

    private void ParseFill(XmlTreeReader xml)
    {
        xml.Skip();
    }
    private void ParseCellAlignment(XmlTreeReader xml)
    {
        xml.Skip();
    }

    private void ParseBorder(XmlTreeReader xml)
    {
        xml.Skip();
    }

    private void ParseCellProtection(XmlTreeReader xml)
    {
        xml.Skip();
    }
}
