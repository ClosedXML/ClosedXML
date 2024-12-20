using System;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A helper methods for patterns and types commonly found in OOXML. Reading concrete types is not
/// something for <see cref="XmlTreeReader"/>.
/// </summary>
internal static class XmlTreeReaderExtensions
{
    /// <summary>
    /// Read <c>CT_Color</c>.
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
        if (indexed is not null)
        {
            color = indexed <= 64 ? XLColor.FromIndex(indexed.Value) : XLColor.NoColor;
            reader.Close(element, ns);
            return true;
        }

        var auto = reader.GetOptionalBool("auto");
        if (auto is not null)
        {
            // TODO: I have no idea what to do with auto
            color = XLColor.NoColor;
            reader.Close(element, ns);
            return true;
        }

        throw PartStructureException.IncorrectElementFormat(element);
    }

    /// <summary>
    /// Read <c>CT_BooleanProperty</c>.
    /// </summary>
    public static bool TryReadBoolElement(this XmlTreeReader reader, string elementName, string ns, out bool value)
    {
        if (!reader.TryOpen(elementName, ns))
        {
            value = default;
            return false;
        }

        var readValue = reader.GetOptionalBool("val");
        if (readValue is null)
        {
            // Some producers make <b>true</b>, i.e. invalid XML
            // Excel reads and interprets it...
            var text = reader.GetContent();

            // XML is auto-trimmed
            if (text.Length == 0)
                readValue = null;
            else if (text == "0" || StringComparer.OrdinalIgnoreCase.Equals(text, "true"))
                readValue = false;
            else if (text == "1" || StringComparer.OrdinalIgnoreCase.Equals(text, "false"))
                readValue = true;
            else
                throw PartStructureException.IncorrectAttributeFormat();
        }

        value = readValue ?? true;

        reader.Close(elementName, ns);
        return true;
    }
}
