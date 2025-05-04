using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen;

internal class SchemeTypeMap
{
    /// <summary>
    /// Simple type map. The key is an XML simple name, the value is C# type name (including
    /// namespace). The type is never nullable, nullability is determined by other XML attributes.
    /// </summary>
    private readonly Dictionary<string, string> _simpleTypeMap = new();

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

    internal SchemeTypeMap AddSimpleTypeRequired(string typeName, string methodTemplate, string cSharpTypeName)
    {
        AddSimpleTypeTemplate(typeName, true, methodTemplate, cSharpTypeName);
        return this;
    }

    internal SchemeTypeMap AddSimpleTypeOptional(string typeName, string methodTemplate, string cSharpTypeName)
    {
        AddSimpleTypeTemplate(typeName, false, methodTemplate, cSharpTypeName);
        return this;
    }

    private void AddSimpleTypeTemplate(string typeName, bool isRequired, string methodTemplate, string cSharpTypeName)
    {
        // Simple type can be added multiple types for optional and required templates
        if (!_simpleTypeMap.TryGetValue(typeName, out var existingCSharpType))
        {
            _simpleTypeMap.Add(typeName, cSharpTypeName);
        }
        else
        {
            if (cSharpTypeName != existingCSharpType)
                throw new InvalidOperationException($"The XML type {typeName} should be mapped to {cSharpTypeName}, but is already mapped to {existingCSharpType}.");
        }

        var typeTemplate = isRequired ? _requiredSimpleTypeTemplate : _optionalSimpleTypeTemplate;
        typeTemplate.Add(typeName, methodTemplate);
    }

    internal string GetSimpleTypeName(string typeName)
    {
        return _simpleTypeMap[typeName];
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

    public SchemeTypeMap AddPrimitiveTypes()
    {
        AddSimpleTypeRequired("xsd:unsignedInt", "_reader.GetUInt(\"{0}\")", "uint");
        AddSimpleTypeOptional("xsd:int", "_reader.GetOptionalInt(\"{0}\")", "int");
        AddSimpleTypeRequired("xsd:boolean", "_reader.GetBool(\"{0}\")", "bool");
        AddSimpleTypeOptional("xsd:boolean", "_reader.GetOptionalBool(\"{0}\")", "bool");
        AddSimpleTypeOptional("s:ST_Xstring", "_reader.GetOptionalXString(\"{0}\")", "string");
        AddSimpleTypeRequired("s:ST_Xstring", "_reader.GetXString(\"{0}\")", "string");
        AddSimpleTypeOptional("xsd:unsignedInt", "_reader.GetOptionalUInt(\"{0}\")", "uint");
        AddSimpleTypeRequired("xsd:dateTime", "_reader.GetDateTime(\"{0}\")", "System.DateTime");
        AddSimpleTypeOptional("ST_UnsignedIntHex", "_reader.GetOptionalUIntHex(\"{0}\")", "uint");
        AddSimpleTypeRequired("xsd:double", "_reader.GetDouble(\"{0}\")", "double");
        return this;
    }
}
