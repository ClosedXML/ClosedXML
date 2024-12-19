using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A universal mapper of string representation of an enum value in the OOXML to ClosedXML enum.
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
        
        return enumMaps;
    }
}
