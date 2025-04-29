using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen;

internal class SchemeTypeMap
{
    /// <summary>
    /// Simple type map. The key is a simple name, while value is C# type. The type is never
    /// nullable, nullability is determined by other XML attributes.
    /// </summary>
    private readonly Dictionary<string, Type> _typeMap = new();

    /// <summary>
    /// Templates for required attributes in XML. The key is XML simple type name, the value is a
    /// code expression template used to get the value. The template has one argument (<c>{0}</c>)
    /// that is substituted with a name of an attribute.
    /// </summary>
    private readonly Dictionary<string, string> _requiredSimpleTypeTemplate = new();

    /// <summary>
    /// Templates for optional attributes in XML. The key is XML simple type name, the value is a
    /// code expression template used to get the value. The template has one argument (<c>{0}</c>)
    /// that is substituted with a name of an attribute.
    /// </summary>
    private readonly Dictionary<string, string> _optionalSimpleTypeTemplate = new();

    internal void AddSimpleTypeTemplate<CSharpType>(string typeName, bool isRequired, string methodTemplate)
    {
        RegisterTypeMapping<CSharpType>(typeName, isRequired);

        var typeTemplate = isRequired ? _requiredSimpleTypeTemplate : _optionalSimpleTypeTemplate;
        typeTemplate.Add(typeName, methodTemplate);
    }

    private void RegisterTypeMapping<CSharpType>(string typeName, bool isRequired)
    {
        var registeredCSharpType = typeof(CSharpType);
        if (!isRequired && registeredCSharpType.IsGenericType && registeredCSharpType.GetGenericTypeDefinition() == typeof(Nullable<>))
            registeredCSharpType = registeredCSharpType.GetGenericArguments()[0];

        if (!_typeMap.TryGetValue(typeName, out var existingCSharpType))
        {
            _typeMap.TryAdd(typeName, registeredCSharpType);
        }
        else
        {
            if (registeredCSharpType != existingCSharpType)
                throw new InvalidOperationException($"Adding XML type {typeName} should be mapped to {typeof(CSharpType)}, but is already mapped to {existingCSharpType}.");
        }
    }

    internal string GetSimpleTypeName(string typeName)
    {
        Type cSharpType = GetSimpleType(typeName);
        return Type.GetTypeCode(cSharpType) switch
        {
            TypeCode.Boolean => "bool",
            TypeCode.Int32 => "int",
            TypeCode.UInt32 => "uint",
            TypeCode.Double => "double",
            TypeCode.String => "string",
            _ => cSharpType.FullName ?? throw new InvalidOperationException("Missing full name")
        };
    }

    internal Type GetSimpleType(string typeName)
    {
        return _typeMap[typeName];
    }

    internal string GetSimpleTypeMethod(AttributeElement attribute)
    {
        var isOptional = attribute.Use is AttributeUseType.Default or AttributeUseType.Optional;
        var templates = isOptional ? _optionalSimpleTypeTemplate : _requiredSimpleTypeTemplate;
        var typeName = attribute.Type ?? throw new InvalidOperationException();
        if (!templates.TryGetValue(typeName, out var methodTemplate))
            throw new InvalidOperationException($"Simple type {typeName} ({(isOptional ? "optional" : "required")}) doesn't have defined template.");

        return string.Format(methodTemplate, attribute.Name);
    }
}
