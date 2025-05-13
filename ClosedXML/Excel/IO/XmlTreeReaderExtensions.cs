using System;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.IO;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A helper methods for patterns and types commonly found in OOXML. Reading concrete types is not
/// something for <see cref="XmlTreeReader"/>.
/// </summary>
internal static class XmlTreeReaderExtensions
{
    /// <summary>
    /// Try to open an optional <c>CT_Color</c> and read the color.
    /// </summary>
    /// <returns><c>true</c> when there was a color element, false if there wasn't.</returns>
    public static bool TryReadColor(this XmlTreeReader reader, string colorElementName, string ns, [NotNullWhen(true)] out XLColor? color)
    {
        if (reader.TryOpen(colorElementName, ns))
        {
            color = reader.ParseColor(colorElementName, ns);
            return true;
        }

        color = default;
        return false;
    }

    /// <summary>
    /// Read <c>CT_Color</c>.
    /// </summary>
    public static XLColor ParseColor(this XmlTreeReader reader, string colorElementName, string ns)
    {
        // OI-29500: Office prioritizes the attributes as auto < indexed < rgb < theme, and only
        // round trips the type with the highest priority if two or more are specified.
        var theme = reader.GetOptionalUInt("theme");
        if (theme is not null)
        {
            var tint = reader.GetOptionalDouble("tint") ?? 0;
            var themeColor = XLColor.FromTheme((XLThemeColor)theme.Value, tint);
            reader.Close(colorElementName, ns);
            return themeColor;
        }

        var rgb = reader.GetOptionalString("rgb");
        if (rgb is not null)
        {
            var rgbColor = XLColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
            reader.Close(colorElementName, ns);
            return rgbColor;
        }

        var indexed = reader.GetOptionalUintAsInt("indexed");
        if (indexed is not null)
        {
            var indexedColor = indexed <= 64 ? XLColor.FromIndex(indexed.Value) : XLColor.NoColor;
            reader.Close(colorElementName, ns);
            return indexedColor;
        }

        var auto = reader.GetOptionalBool("auto");
        if (auto is not null)
        {
            // TODO: I have no idea what to do with auto
            var autoColor = XLColor.NoColor;
            reader.Close(colorElementName, ns);
            return autoColor;
        }

        throw PartStructureException.IncorrectElementFormat(colorElementName);
    }

    /// <summary>
    /// Read <c>CT_BooleanProperty</c>.
    /// </summary>
    public static bool TryReadBoolElement(this XmlTreeReader reader, string boolElementName, string ns, out bool value)
    {
        if (!reader.TryOpen(boolElementName, ns))
        {
            value = default;
            return false;
        }

        value = reader.GetOptionalBool("val") ?? true;

        reader.Close(boolElementName, ns);
        return true;
    }

    /// <summary>
    /// Try to read an element with an attribute <c>val</c> that contains a <c>XString</c>. Example:
    /// <![CDATA[
    /// <xsd:complexType name="CT_FontName">
    ///   <xsd:attribute name="val" type="s:ST_Xstring" use="required"/>
    /// </xsd:complexType>
    /// ]]>
    /// </summary>
    public static bool TryReadXStringValElement(this XmlTreeReader reader, string elementName, string ns, [NotNullWhen(true)] out string? value)
    {
        if (reader.TryOpen(elementName, ns))
        {
            value = reader.GetXString("val");
            reader.Close(elementName, ns);
            return true;
        }

        value = null;
        return false;
    }

    /// <summary>
    /// Try to read an optional element of type <c>CT_IntProperty</c>
    /// </summary>
    public static bool TryReadIntValElement(this XmlTreeReader reader, string elementName, string ns, out int value)
    {
        if (reader.TryOpen(elementName, ns))
        {
            value = reader.GetInt("val");
            reader.Close(elementName, ns);
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to read an optional element that contains a required enum in a <c>val</c> attribute.
    /// <![CDATA[
    /// <xsd:complexType name="CT_FontScheme">
    ///   <xsd:attribute name="val" type="ST_FontScheme" use="required"/>
    /// </xsd:complexType>
    /// ]]>
    /// </summary>
    /// <returns><c>true</c> when there was the element, <c>false</c> when there wasn't the element.</returns>
    public static bool TryReadEnumValElement<TEnum>(this XmlTreeReader reader, string elementName, string ns, [NotNullWhen(true)] out TEnum? enumValue)
        where TEnum : struct, Enum
    {
        if (reader.TryOpen(elementName, ns))
        {
            enumValue = reader.GetEnum<TEnum>("val");
            reader.Close(elementName, ns);
            return true;
        }

        enumValue = null;
        return false;
    }
}
