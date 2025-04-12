using System;
using System.Collections.Generic;
using ClosedXML.IO;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A universal two-way mapper of string representation of an enum value in the OOXML to ClosedXML enum.
/// </summary>
internal sealed class XmlToEnumMapper : IEnumMapper
{
    /// <summary>
    /// A collection of all maps. The key is enum type, the value is Dictionary&lt;string,SomeEnum&gt;
    /// Value can't be typed due to generic limitations (no common ancestor).
    /// </summary>
    private readonly Dictionary<Type, object> _textToEnumMaps;

    private static readonly Lazy<XmlToEnumMapper> LazyInstance = new(CreateSpreadsheetMapper);

    internal static XmlToEnumMapper Instance => LazyInstance.Value;

    private XmlToEnumMapper(Dictionary<Type, object> maps)
    {
        _textToEnumMaps = maps;
    }

    public bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        var enumMap = (Dictionary<string, TEnum>)_textToEnumMaps[typeof(TEnum)];
        return enumMap.TryGetValue(text, out enumValue);
    }

    private static XmlToEnumMapper CreateSpreadsheetMapper()
    {
        var builder = new Builder();

        // ST_FontScheme
        builder.Add(new Dictionary<string, XLFontScheme>
        {
            { "none", XLFontScheme.None },
            { "major", XLFontScheme.Major },
            { "minor", XLFontScheme.Minor },
        });

        // ST_UnderlineValues
        builder.Add(new Dictionary<string, XLFontUnderlineValues>
        {
            { "double", XLFontUnderlineValues.Double },
            { "doubleAccounting", XLFontUnderlineValues.DoubleAccounting },
            { "none", XLFontUnderlineValues.None },
            { "single", XLFontUnderlineValues.Single },
            { "singleAccounting", XLFontUnderlineValues.SingleAccounting },
        });

        // ST_VerticalAlignRun
        builder.Add(new Dictionary<string, XLFontVerticalTextAlignmentValues>
        {
            { "baseline", XLFontVerticalTextAlignmentValues.Baseline },
            { "subscript", XLFontVerticalTextAlignmentValues.Subscript },
            { "superscript", XLFontVerticalTextAlignmentValues.Superscript },
        });

        return builder.Build();
    }

    internal class Builder
    {
        private readonly Dictionary<Type, object> _maps = new();

        public Builder Add<T>(Dictionary<string, T> map)
        {
            _maps.Add(typeof(T), map);
            return this;
        }

        public XmlToEnumMapper Build()
        {
            return new XmlToEnumMapper(_maps);
        }
    }
}
