using System;
using System.Collections.Generic;
using ClosedXML.IO;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A universal two-way mapper of string representation of an enum value in the OOXML to ClosedXML enum.
/// </summary>
internal sealed class XmlToEnumMapper : IEnumMapper
{
    /// <summary>
    /// A collection of all maps. The key is enum type, the value is BiDictionary&lt;string,SomeEnum&gt;
    /// Value can't be typed due to generic limitations (no common ancestor).
    /// </summary>
    private readonly Dictionary<Type, object> _enumMaps;

    private static readonly Lazy<XmlToEnumMapper> LazyInstance = new(CreateSpreadsheetMapper);

    internal static XmlToEnumMapper Instance => LazyInstance.Value;

    private XmlToEnumMapper(Dictionary<Type, object> maps)
    {
        _enumMaps = maps;
    }

    public bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        var enumMap = GetEnumMap<TEnum>();
        return enumMap.KeyToValue.TryGetValue(text, out enumValue);
    }

    public bool TryGetText<TEnum>(TEnum enumValue, out string text)
        where TEnum : struct, Enum
    {
        var enumMap = GetEnumMap<TEnum>();
        return enumMap.ValueToKey.TryGetValue(enumValue, out text);
    }

    private BiDictionary<string, TEnum> GetEnumMap<TEnum>()
        where TEnum : struct, Enum
    {
        return (BiDictionary<string, TEnum>)_enumMaps[typeof(TEnum)];
    }

    private static XmlToEnumMapper CreateSpreadsheetMapper()
    {
        var builder = new Builder();

        // ST_FontScheme
        builder.Add(new BiDictionary<string, XLFontScheme>
        {
            { "none", XLFontScheme.None },
            { "major", XLFontScheme.Major },
            { "minor", XLFontScheme.Minor },
        });

        // ST_UnderlineValues
        builder.Add(new BiDictionary<string, XLFontUnderlineValues>
        {
            { "double", XLFontUnderlineValues.Double },
            { "doubleAccounting", XLFontUnderlineValues.DoubleAccounting },
            { "none", XLFontUnderlineValues.None },
            { "single", XLFontUnderlineValues.Single },
            { "singleAccounting", XLFontUnderlineValues.SingleAccounting },
        });

        // ST_VerticalAlignRun
        builder.Add(new BiDictionary<string, XLFontVerticalTextAlignmentValues>
        {
            { "baseline", XLFontVerticalTextAlignmentValues.Baseline },
            { "subscript", XLFontVerticalTextAlignmentValues.Subscript },
            { "superscript", XLFontVerticalTextAlignmentValues.Superscript },
        });

        // ST_PatternType
        builder.Add(new BiDictionary<string, XLFillPatternValues>
        {
            { "none", XLFillPatternValues.None },
            { "solid", XLFillPatternValues.Solid },
            { "mediumGray", XLFillPatternValues.MediumGray },
            { "darkGray", XLFillPatternValues.DarkGray },
            { "lightGray", XLFillPatternValues.LightGray },
            { "darkHorizontal", XLFillPatternValues.DarkHorizontal },
            { "darkVertical", XLFillPatternValues.DarkVertical },
            { "darkDown", XLFillPatternValues.DarkDown },
            { "darkUp", XLFillPatternValues.DarkUp },
            { "darkGrid", XLFillPatternValues.DarkGrid },
            { "darkTrellis", XLFillPatternValues.DarkTrellis },
            { "lightHorizontal", XLFillPatternValues.LightHorizontal },
            { "lightVertical", XLFillPatternValues.LightVertical },
            { "lightDown", XLFillPatternValues.LightDown },
            { "lightUp", XLFillPatternValues.LightUp },
            { "lightGrid", XLFillPatternValues.LightGrid },
            { "lightTrellis", XLFillPatternValues.LightTrellis },
            { "gray125", XLFillPatternValues.Gray125 },
            { "gray0625", XLFillPatternValues.Gray0625 },
        });

        // ST_BorderStyle
        builder.Add(new BiDictionary<string, XLBorderStyleValues>
        {
            { "none", XLBorderStyleValues.None },
            { "thin", XLBorderStyleValues.Thin },
            { "medium", XLBorderStyleValues.Medium },
            { "dashed", XLBorderStyleValues.Dashed },
            { "dotted", XLBorderStyleValues.Dotted },
            { "thick", XLBorderStyleValues.Thick },
            { "double", XLBorderStyleValues.Double },
            { "hair", XLBorderStyleValues.Hair },
            { "mediumDashed", XLBorderStyleValues.MediumDashed },
            { "dashDot", XLBorderStyleValues.DashDot },
            { "mediumDashDot", XLBorderStyleValues.MediumDashDot },
            { "dashDotDot", XLBorderStyleValues.DashDotDot },
            { "mediumDashDotDot", XLBorderStyleValues.MediumDashDotDot },
            { "slantDashDot", XLBorderStyleValues.SlantDashDot },
        });

        // ST_HorizontalAlignment
        builder.Add(new BiDictionary<string, XLAlignmentHorizontalValues>
        {
            { "general", XLAlignmentHorizontalValues.General },
            { "left", XLAlignmentHorizontalValues.Left },
            { "center", XLAlignmentHorizontalValues.Center },
            { "right", XLAlignmentHorizontalValues.Right },
            { "fill", XLAlignmentHorizontalValues.Fill },
            { "justify", XLAlignmentHorizontalValues.Justify },
            { "centerContinuous", XLAlignmentHorizontalValues.CenterContinuous },
            { "distributed", XLAlignmentHorizontalValues.Distributed },
        });

        // ST_VerticalAlignment
        builder.Add(new BiDictionary<string, XLAlignmentVerticalValues>
        {
            { "top", XLAlignmentVerticalValues.Top },
            { "center", XLAlignmentVerticalValues.Center },
            { "bottom", XLAlignmentVerticalValues.Bottom },
            { "justify", XLAlignmentVerticalValues.Justify },
            { "distributed", XLAlignmentVerticalValues.Distributed },
        });

        // ST_GradientType
        builder.Add(new BiDictionary<string, XLGradientType>
        {
            { "linear", XLGradientType.Linear },
            { "path", XLGradientType.Path }
        });

        return builder.Build();
    }

    internal class Builder
    {
        private readonly Dictionary<Type, object> _maps = new();

        public Builder Add<T>(BiDictionary<string, T> map)
            where T : struct, Enum
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
