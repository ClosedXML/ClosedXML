using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using ClosedXML.Excel.Formatting;
using ClosedXML.IO;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;

namespace ClosedXML.Excel.IO;

internal class StylesPartReader
{
    internal static XLWorkbookStyles Load(Stylesheet stylesheet)
    {
        // Load fonts
        var xlFontFormats = new List<XLFontFormat>();
        if (stylesheet.Fonts is { } fonts)
        {
            xlFontFormats = LoadFonts(fonts);
        }

        // Load master formatting records, generally used by cells, but also by rows or columns
        var xlCellFormats = new List<XLCellFormat>();
        if (stylesheet.CellFormats is { } cellFormats)
        {
            xlCellFormats = LoadCellFormats(cellFormats, xlFontFormats);
        }

        return new XLWorkbookStyles(xlCellFormats, xlFontFormats);
    }

    private static List<XLFontFormat> LoadFonts(Fonts fonts)
    {
        var xlFontFormats = new List<XLFontFormat>();
        foreach (var font in fonts.Elements<Font>())
        {
            var xlFontFormat = LoadFont(font);
            xlFontFormats.Add(xlFontFormat);
        }

        return xlFontFormats;
    }

    private static XLFontFormat LoadFont(Font font)
    {
        var name = font.FontName?.Val?.Value;
        var charset = (XLFontCharSet?)LoadOptionalIntElement(font.FontCharSet);
        var family = LoadIntEnumElement<XLFontFamilyNumberingValues>(font.FontFamilyNumbering);
        var bold = LoadOptionalBoolElement(font.Bold);
        var italic = LoadOptionalBoolElement(font.Italic);
        var strikethrough = LoadOptionalBoolElement(font.Strike);
        var outline = LoadOptionalBoolElement(font.Outline);
        var shadow = LoadOptionalBoolElement(font.Shadow);
        var condense = LoadOptionalBoolElement(font.Condense);
        var extend = LoadOptionalBoolElement(font.Extend);
        var color = LoadOptionalColorElement(font.Color);
        var sizePt = LoadOptionalDoubleElement(font.FontSize);
        var underline = LoadOptionalEnumElement(font.Underline, XLFontUnderlineValues.Single);
        var verticalAlignment = LoadOptionalEnumElement<XLFontVerticalTextAlignmentValues>(font.VerticalTextAlignment);
        var scheme = LoadOptionalEnumElement<XLFontScheme>(font.FontScheme);

        return new XLFontFormat
        {
            Name = name,
            Charset = charset,
            Family = family,
            Bold = bold,
            Italic = italic,
            Strikethrough = strikethrough,
            Outline = outline,
            Shadow = shadow,
            Condense = condense,
            Extend = extend,
            Color = color,
            Size = XLFontSize.FromPoints(sizePt),
            Underline = underline,
            VerticalAlignment = verticalAlignment,
            Scheme = scheme,
        };
    }

    /// <inheritdoc cref="LoadEnumElement{TEnum}(OpenXmlLeafElement,TEnum)"/>
    /// <remarks>
    /// Element is referenced as an optional.
    /// <code><![CDATA[
    ///   <xsd:element name="elementName" type="ST_SomeEnum" minOccurs="0" maxOccurs="1"/>
    /// ]]></code>
    /// </remarks>
    private static TEnum? LoadOptionalEnumElement<TEnum>(OpenXmlLeafElement? element, TEnum defaultValue)
        where TEnum : struct, Enum
    {
        if (element is null)
            return null;

        return LoadEnumElement(element, defaultValue);
    }

    /// <summary>
    /// Parse enum with a default value.
    /// <code><![CDATA[
    /// <xsd:complexType name="CT_SomeEnumType">
    ///   <xsd:attribute name="val" type="ST_SomeEnum" use="optional" default="foo"/>
    /// </xsd:complexType>
    /// 
    /// <xsd:simpleType name="ST_SomeEnum">
    ///   <xsd:restriction base="xsd:string">
    ///     <xsd:enumeration value = "foo"/>
    ///     <xsd:enumeration value = "bar"/>
    ///   </xsd:restriction>
    /// </xsd:simpleType>
    /// ]]></code>
    /// </summary>
    private static TEnum LoadEnumElement<TEnum>(OpenXmlLeafElement element, TEnum defaultValue)
        where TEnum : struct, Enum
    {
        // OpenXML SDK doesn't have a proper way to check for a specific attribute
        if (!element.HasAttributes || element.GetAttributes().All(e => e.LocalName != "val"))
            return defaultValue;

        var attributeText = element.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(attributeText))
            throw PartStructureException.InvalidAttributeFormat();

        if (!XmlToEnumMapper.Instance.TryGetEnum<TEnum>(attributeText, out var enumValue))
            throw PartStructureException.InvalidAttributeFormat(attributeText);

        return enumValue;
    }

    /// <inheritdoc cref="LoadEnumElement{TEnum}(OpenXmlLeafElement)"/>
    /// <remarks>
    /// Element is referenced as an optional.
    /// <code><![CDATA[
    ///   <xsd:element name="elementName" type="ST_SomeEnum" minOccurs="0" maxOccurs="1"/>
    /// ]]></code>
    /// </remarks>
    private static TEnum? LoadOptionalEnumElement<TEnum>(OpenXmlLeafElement? element)
        where TEnum : struct, Enum
    {
        if (element is null)
            return null;

        return LoadEnumElement<TEnum>(element);
    }

    /// <summary>
    /// Parse enum.
    /// <code><![CDATA[
    /// <xsd:complexType name="CT_SomeEnumType">
    ///   <xsd:attribute name="val" type="ST_SomeEnum" use="required"/>
    /// </xsd:complexType>
    /// 
    /// <xsd:simpleType name="ST_SomeEnum">
    ///   <xsd:restriction base="xsd:string">
    ///     <xsd:enumeration value = "foo"/>
    ///     <xsd:enumeration value = "bar"/>
    ///   </xsd:restriction>
    /// </xsd:simpleType>
    /// ]]></code>
    /// </summary>
    private static TEnum LoadEnumElement<TEnum>(OpenXmlLeafElement element)
        where TEnum : struct, Enum
    {
        var attributeText = element.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(attributeText))
            throw PartStructureException.InvalidAttributeFormat();

        if (!XmlToEnumMapper.Instance.TryGetEnum<TEnum>(attributeText, out var enumValue))
            throw PartStructureException.InvalidAttributeFormat(attributeText);

        return enumValue;
    }

    /// <summary>
    /// Load an enum stored in an element. The value is stored as an integer. E.g.,
    /// <c>&lt;someEnumElement val="12" /&gt;</c>. This method assumes that <typeparamref name="TEnum"/>
    /// values correspond to the int values in the elements attribute.
    /// <code><![CDATA[
    /// <xsd:complexType name="CT_SomeEnum">
    ///   <xsd:attribute name = "val" type="ST_SomeEnumValues" use="required"/>
    /// </xsd:complexType>
    /// 
    /// <xsd:simpleType name = "ST_SomeEnumValues">
    ///   <xsd:restriction base="xsd:integer">
    ///     <xsd:minInclusive value="0"/>
    ///     <xsd:maxInclusive value="14"/>
    ///   </xsd:restriction>
    /// </xsd:simpleType>
    /// ]]></code>
    /// </summary>
    /// <param name="element">The element to parse.</param>
    private static TEnum? LoadIntEnumElement<TEnum>(OpenXmlLeafElement? element)
        where TEnum : struct, Enum
    {
        if (element is null)
            return null;

        var attributeText = element.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(attributeText))
            throw PartStructureException.InvalidAttributeFormat();

        var intValue = XmlConvert.ToInt32(attributeText);
        if (!Enum.IsDefined(typeof(TEnum), intValue))
            throw PartStructureException.InvalidAttributeFormat(attributeText);

        return (TEnum)Enum.ToObject(typeof(TEnum), intValue);
    }

    private static List<XLCellFormat> LoadCellFormats(CellFormats cellFormats, IReadOnlyList<XLFontFormat> xlFontFormats)
    {
        var xlCellFormats = new List<XLCellFormat>();
        foreach (var cellFormat in cellFormats.Elements<CellFormat>())
        {
            var xlCellFormat = LoadCellFormat(cellFormat, xlFontFormats);
            xlCellFormats.Add(xlCellFormat);
        }

        return xlCellFormats;
    }

    private static XLCellFormat LoadCellFormat(CellFormat cellFormat, IReadOnlyList<XLFontFormat> xlFontFormats)
    {
        var xlCellFormat = new XLCellFormat();
        if (cellFormat.FontId?.Value is { } fontId && fontId < xlFontFormats.Count)
        {
            xlCellFormat = new XLCellFormat { Font = xlFontFormats[(int)fontId] };
        }

        return xlCellFormat;
    }

    /// <summary>
    /// Read optional <c>CT_BooleanProperty</c>.
    /// </summary>
    private static bool? LoadOptionalBoolElement(BooleanPropertyType? optionalElement)
    {
        return optionalElement is not null ? LoadBoolElement(optionalElement) : null;
    }

    /// <summary>
    /// Read <c>CT_BooleanProperty</c>.
    /// </summary>
    private static bool LoadBoolElement(BooleanPropertyType element)
    {
        return element.Val?.Value ?? true; // Default value is true
    }

    /// <summary>
    /// Read optional <c>CT_IntProperty</c>.
    /// </summary>
    private static int? LoadOptionalIntElement(OpenXmlLeafElement? optionalElement)
    {
        return optionalElement is not null ? LoadIntElement(optionalElement) : null;
    }

    /// <summary>
    /// Read <c>CT_IntProperty</c>.
    /// </summary>
    private static int LoadIntElement(OpenXmlLeafElement element)
    {
        var attributeText = element.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(attributeText))
            throw PartStructureException.InvalidAttributeFormat();

        return XmlConvert.ToInt32(attributeText);
    }

    /// <inheritdoc cref="LoadDoubleElement"/>
    private static double? LoadOptionalDoubleElement(OpenXmlLeafElement? optionalElement)
    {
        return optionalElement is not null ? LoadDoubleElement(optionalElement) : null;
    }

    /// <summary>
    /// Load a double stored in an element. E.g., <c>&lt;someEnumElement val="12.4" /&gt;</c>. The
    /// <c>val</c> is a required attribute. The schema doesn't have directly an element for that,
    /// but there are several compatible ones like <c>CT_FontSize</c>.
    /// <code><![CDATA[
    /// <xsd:complexType name="CT_SomeType">
    ///   <xsd:attribute name="val" type="xsd:double" use="required"/>
    /// </xsd:complexType>
    /// ]]></code>
    /// </summary>
    private static double LoadDoubleElement(OpenXmlLeafElement element)
    {
        var attributeText = element.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(attributeText))
            throw PartStructureException.InvalidAttributeFormat();

        return XmlConvert.ToDouble(attributeText);
    }

    /// <summary>
    /// Load an <c>CT_Color</c> from an optional element. It should be declared as <code><![CDATA[
    /// <xsd:element name="colorElementName" type="CT_Color" minOccurs="0" maxOccurs="1"/>
    /// ]]></code>
    /// </summary>
    private static XLColor? LoadOptionalColorElement(Color? optionalElement)
    {
        return optionalElement is not null ? LoadColorElement(optionalElement) : null;
    }

    /// <summary>
    /// Load an <c>CT_Color</c> from an required element. It should be declared as <code><![CDATA[
    /// <xsd:element name="colorElementName" type="CT_Color" />
    /// ]]></code>
    /// </summary>
    private static XLColor LoadColorElement(Color element)
    {
        return element.ToClosedXMLColor();
    }
}
