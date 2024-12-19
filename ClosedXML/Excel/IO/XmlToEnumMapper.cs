using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A universal two-way mapper of string representation of an enum value in the OOXML to ClosedXML enum.
/// </summary>
internal static class XmlToEnumMapper
{
    /// <summary>
    /// A collection of all maps. The key is enum type, the value is Dictionary&lt;string,SomeEnum&gt;
    /// Value can't be typed due to generic limitations (no common ancestor).
    /// </summary>
    private static readonly Lazy<Dictionary<Type, object>> TextToEnumMaps = new(CreateMaps);

    public static bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        var enumMap = (Dictionary<string, TEnum>)TextToEnumMaps.Value[typeof(TEnum)];
        return enumMap.TryGetValue(text, out enumValue);
    }

    private static Dictionary<Type, object> CreateMaps()
    {
        var enumMaps = new Dictionary<Type, object>();

        // ST_FontScheme
        var xlFontScheme = new Dictionary<string, XLFontScheme>
        {
            { "none", XLFontScheme.None },
            { "major", XLFontScheme.Major },
            { "minor", XLFontScheme.Minor },
        };
        enumMaps.Add(typeof(XLFontScheme), xlFontScheme);

        // ST_UnderlineValues
        var xlFontUnderline = new Dictionary<string, XLFontUnderlineValues>
        {
            { "double", XLFontUnderlineValues.Double },
            { "doubleAccounting", XLFontUnderlineValues.DoubleAccounting },
            { "none", XLFontUnderlineValues.None },
            { "single", XLFontUnderlineValues.Single },
            { "singleAccounting", XLFontUnderlineValues.SingleAccounting },
        };
        enumMaps.Add(typeof(XLFontUnderlineValues), xlFontUnderline);

        // ST_VerticalAlignRun
        var xlFontVerticalTextAlignmentValues = new Dictionary<string, XLFontVerticalTextAlignmentValues>
        {
            {"baseline",XLFontVerticalTextAlignmentValues.Baseline},
            {"subscript",XLFontVerticalTextAlignmentValues.Subscript },
            {"superscript",XLFontVerticalTextAlignmentValues.Superscript },
        };
        enumMaps.Add(typeof(XLFontVerticalTextAlignmentValues), xlFontVerticalTextAlignmentValues);

        return enumMaps;
    }
}
