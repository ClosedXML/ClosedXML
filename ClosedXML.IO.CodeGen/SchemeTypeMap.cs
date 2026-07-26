using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.IO.CodeGen;

public class SchemeTypeMap
{
    /// <summary>
    /// Simple type map. The key is an XML simple name, the value is info about how to work with it in the C# code.
    /// </summary>
    private readonly Dictionary<string, SimpleTypeMapping> _simpleTypeMap = new();

    /// <summary>
    /// Map of XML complex type or element group name to C# type (as a text). If there isn't a record in the map,
    /// there is no mapping codegen will use <c>void</c>.
    /// </summary>
    private readonly Dictionary<ParsletName, (string FinalType, string ItemType)> _parsletMap = new();

    /// <summary>
    /// A dictionary of parse calls, i.e. a code used to call a parslet. Under normal circumstance,
    /// they are are automatically generated for TLE (<c>CT_Shape</c> - <c>ParseShape</c>), but
    /// there are two useful scenarios for manual specification:
    /// <list type="bullet">
    ///   <item>Different schema - That is likely specified in a separate reader that is a field of
    ///     current reader. So it should generate something like `_colorReader.ParseColor()`</item>
    ///   <item>Method name collision - all parslets are private, but reusable readers need
    ///     to expose internal methods. Since they often have same name, it could be useful to avoid
    ///     collision.</item>
    /// </list>
    /// </summary>
    private readonly Dictionary<ParsletName, string> _manualParseCall = new();

    public SchemeTypeMap AddComplexTypeMapping(ParsletName complexTypeName, string cSharpType)
    {
        _parsletMap.Add(complexTypeName, (cSharpType, cSharpType));
        return this;
    }

    public SchemeTypeMap AddComplexTypeMapping(ParsletName complexTypeName, string cSharpType, string csItemType)
    {
        _parsletMap.Add(complexTypeName, (cSharpType, csItemType));
        return this;
    }

    public SchemeTypeMap AddSimpleType(SimpleTypeMapping simpleType)
    {
        _simpleTypeMap.Add(simpleType.Name, simpleType);
        return this;
    }

    public SchemeTypeMap AddSimpleTypeEnum(string simpleType, string csTypeName, string xmlValue, string csValue)
    {
        return AddSimpleTypeEnum(simpleType, csTypeName, new() { { xmlValue, csValue } });
    }

    public SchemeTypeMap AddSimpleTypeEnum(string simpleType, string csTypeName, Dictionary<string, string>? valuesMap = null)
    {
        return AddSimpleType(new SimpleTypeMapping
        {
            Name = simpleType,
            CsTypeName = csTypeName,
            RequiredTemplate = $"_reader.GetEnum<{csTypeName}>(\"{{0}}\")",
            OptionalTemplate = $"_reader.GetOptionalEnum<{csTypeName}>(\"{{0}}\")",
            MapValue = xmlName => valuesMap?[xmlName] ?? throw new InvalidOperationException($"The XML value {xmlName} is not mapped to {csTypeName}.")
        });
    }

    /// <summary>
    /// Specify a piece of code that will be used to parse <paramref name="name"/> type.  It must
    /// return the correct type that was specified by the <see cref="AddComplexTypeMapping(ParsletName,string)"/>
    /// </summary>
    /// <param name="name">Parslet for which the call is defined.</param>
    /// <param name="parseCall">A name of method or a call of another field without parenthesis, e.g. <c>_reader.ParsePoint</c>.</param>
    /// <returns></returns>
    public SchemeTypeMap AddParseCall(ParsletName name, string parseCall)
    {
        _manualParseCall.Add(name, parseCall);
        return this;
    }

    internal SimpleTypeMapping GetSimpleType(string typeName)
    {
        return _simpleTypeMap[typeName];
    }

    internal string GetSimpleTypeMethod(AttributeElement attribute)
    {
        var simpleTypeName = attribute.Type ?? throw new InvalidOperationException();
        var simpleType = _simpleTypeMap[simpleTypeName];
        var expressionTemplate = attribute.IsOptional ? simpleType.OptionalTemplate : simpleType.RequiredTemplate;
        return string.Format(expressionTemplate, attribute.Name);
    }

    internal bool TryGetParsletCsType(ParsletName name, [NotNullWhen(true)] out string? csType)
    {
        if (_parsletMap.TryGetValue(name, out var map))
        {
            csType = map.FinalType;
            return true;
        }

        csType = null;
        return false;
    }

    internal bool TryItemGetValue(ParsletName parsletName, [NotNullWhen(true)] out string? csItemType)
    {
        if (_parsletMap.TryGetValue(parsletName, out var map))
        {
            csItemType = map.ItemType;
            return true;
        }

        csItemType = null;
        return false;
    }

    internal string GetParseCall(ParsletName name)
    {
        if (_manualParseCall.TryGetValue(name, out var manualParseCall))
            return manualParseCall;

        if (name.HasNamespace)
            throw new InvalidOperationException($"Parslet '{name}' uses a namespace. Specify the parseCall manually through the '{nameof(SchemeTypeMap)}.{nameof(AddParseCall)}()' method");

        return $"Parse{name.WithoutPrefix()}";
    }

    public SchemeTypeMap AddPrimitiveTypes()
    {
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:boolean",
            CsTypeName = "bool",
            RequiredTemplate = "_reader.GetBool(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalBool(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:byte",
            CsTypeName = "byte",
            RequiredTemplate = "_reader.GetByte(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalByte(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:int",
            CsTypeName = "int",
            RequiredTemplate = "_reader.GetInt(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalInt(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:unsignedInt",
            CsTypeName = "uint",
            RequiredTemplate = "_reader.GetUInt(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:double",
            CsTypeName = "double",
            RequiredTemplate = "_reader.GetDouble(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalDouble(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "s:ST_Xstring",
            CsTypeName = "string",
            RequiredTemplate = "_reader.GetXString(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalXString(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:string",
            CsTypeName = "string",
            RequiredTemplate = "_reader.GetString(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalString(\"{0}\")",
            MapValue = x => x.Length == 0 ? "string.Empty" : $"\"{x.Replace("\"", "\\\"")}\""
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "xsd:dateTime",
            CsTypeName = "System.DateTime",
            RequiredTemplate = "_reader.GetDateTime(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalDateTime(\"{0}\")"
        });
        AddSimpleType(new SimpleTypeMapping
        {
            Name = "ST_UnsignedIntHex",
            CsTypeName = "uint",
            OptionalTemplate = "_reader.GetOptionalUIntHex(\"{0}\")"
        });
        return this;
    }
}
