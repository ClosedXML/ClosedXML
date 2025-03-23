using System;
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
    /// Read <c>CT_Color</c>.
    /// </summary>
    public static bool TryParseColor(this XmlTreeReader reader, string colorElementName, string ns, out XLColor color)
    {
        if (!reader.TryOpen(colorElementName, ns))
        {
            color = XLColor.NoColor;
            return false;
        }

        // OI-29500: Office prioritizes the attributes as auto < indexed < rgb < theme, and only
        // round trips the type with the highest priority if two or more are specified.
        var theme = reader.GetOptionalUint("theme");
        if (theme is not null)
        {
            var tint = reader.GetOptionalDouble("theme") ?? 0;
            color = XLColor.FromTheme((XLThemeColor)theme.Value, tint);
            reader.Close(colorElementName, ns);
            return true;
        }

        var rgb = reader.GetOptionalString("rgb");
        if (rgb is not null)
        {
            color = XLColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
            reader.Close(colorElementName, ns);
            return true;
        }

        var indexed = reader.GetOptionalUintAsInt("indexed");
        if (indexed is not null)
        {
            color = indexed <= 64 ? XLColor.FromIndex(indexed.Value) : XLColor.NoColor;
            reader.Close(colorElementName, ns);
            return true;
        }

        var auto = reader.GetOptionalBool("auto");
        if (auto is not null)
        {
            // TODO: I have no idea what to do with auto
            color = XLColor.NoColor;
            reader.Close(colorElementName, ns);
            return true;
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
}
