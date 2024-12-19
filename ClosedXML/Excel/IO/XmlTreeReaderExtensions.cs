using System;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A helper methods for patterns and types commonly found in OOXML. Reading concrete types is not something for <see cref="XmlTreeReader"/>.
/// </summary>
internal static class XmlTreeReaderExtensions
{
    /// <summary>
    /// Try to parse <c>CT_Color</c>
    /// </summary>
    public static bool TryParseColor(this XmlTreeReader reader, string element, string ns, out XLColor color)
    {
        if (!reader.TryOpen(element, ns))
        {
            color = XLColor.NoColor;
            return false;
        }

        // OI-29500: Office prioritizes the attributes as auto < indexed < rgb < theme, and only
        // round trips the type with the highest priority if two or more are specified.
        var theme = reader.GetOptionalUint("theme");
        if (theme is not null)
        {
            var tint = reader.GetOptionalDouble("theme", 0);
            color = tint is not null
                ? XLColor.FromTheme((XLThemeColor)theme.Value, tint.Value)
                : XLColor.FromTheme((XLThemeColor)theme.Value);
            reader.Close(element, ns);
            return true;
        }

        var rgb = reader.GetOptionalString("rgb");
        if (rgb is not null)
        {
            color = XLColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
            reader.Close(element, ns);
            return true;
        }

        var indexed = reader.GetOptionalUint("indexed");
        if (indexed <= 64)
        {
            color = XLColor.FromIndex(indexed.Value);
            reader.Close(element, ns);
            return true;
        }

        _ = reader.GetOptionalBool("auto"); // IGNORE: auto attribute.

        throw PartStructureException.IncorrectElementFormat(element);
    }

    /// <summary>
    /// Read <c>CT_BooleanProperty</c>.
    /// </summary>
    public static bool TryReadBoolElement(this XmlTreeReader reader, string elementName, out bool value)
    {
        if (!reader.TryOpen(elementName, OpenXmlConst.Main2006SsNs))
        {
            value = default;
            return false;
        }

        value = reader.GetOptionalBool("val") ?? true;
        reader.Close(elementName, OpenXmlConst.Main2006SsNs);
        return true;
    }
}
