using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.IO.CodeGen;

internal class SchemeTypeMap
{
    /// <summary>
    /// Simple type map. The key is an XML simple name, the value is info about how to work with it in the C# code.
    /// </summary>
    private readonly Dictionary<string, SimpleTypeMapping> _simpleTypeMap = new();

    /// <summary>
    /// Map of XML complex type or element group name to C# type (as a text). If there isn't a record in the map,
    /// there is no mapping codegen will use <c>void</c>.
    /// </summary>
    private readonly Dictionary<ParsletName, string> _parsletMap = new();

    internal SchemeTypeMap AddComplexTypeMapping(ParsletName complexTypeName, string cSharpType)
    {
        _parsletMap.Add(complexTypeName, cSharpType);
        return this;
    }

    internal SchemeTypeMap AddElementGroupMapping(ParsletName elementGroup, string cSharpType)
    {
        _parsletMap.Add(elementGroup, cSharpType);
        return this;
    }

    internal SchemeTypeMap AddSimpleType(SimpleTypeMapping simpleType)
    {
        _simpleTypeMap.Add(simpleType.Name, simpleType);
        return this;
    }

    internal SchemeTypeMap AddSimpleTypeEnum(string simpleType, string csTypeName, string xmlValue, string csValue)
    {
        return AddSimpleTypeEnum(simpleType, csTypeName, new() { { xmlValue, csValue } });
    }

    internal SchemeTypeMap AddSimpleTypeEnum(string simpleType, string csTypeName, Dictionary<string, string>? valuesMap = null)
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
        return _parsletMap.TryGetValue(name, out csType);
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
            OptionalTemplate = "_reader.GetOptionalString(\"{0}\")"
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
