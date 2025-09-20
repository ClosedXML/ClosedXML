using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model;

namespace ClosedXML.IO.CodeGen.XsdParser;

/// <summary>
/// Mapper for enums found in XSD of OOXML.
/// </summary>
public class XsdEnumMapper : IEnumMapper
{
    private readonly Dictionary<Type, object> _textToEnumMaps = new();

    public XsdEnumMapper()
    {
        AddMaps();
    }

    public bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        var enumMap = (Dictionary<string, TEnum>)_textToEnumMaps[typeof(TEnum)];
        return enumMap.TryGetValue(text, out enumValue);
    }

    public bool TryGetText<TEnum>(TEnum enumValue, out string text)
        where TEnum : struct, Enum
    {
        throw new NotSupportedException();
    }

    private void AddMaps()
    {
        AddMap(new Dictionary<string, AttributeUseType>
        {
            { "optional", AttributeUseType.Optional },
            { "required", AttributeUseType.Required }
        });
        AddMap(new Dictionary<string, ProcessContents>
        {
            { "strict", ProcessContents.Strict },
            { "lax", ProcessContents.Lax }
        });
    }

    private void AddMap<TEnum>(Dictionary<string, TEnum> enumMap)
        where TEnum : struct, Enum
    {
        _textToEnumMaps.Add(typeof(TEnum), enumMap);
    }
}
