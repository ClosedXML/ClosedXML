using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.IO.CodeGen;

internal class CodeBuilder
{
    /// <summary>
    /// C# keywords. The variables with that name must be escaped, e.g. <c>in</c> must be <c>@in</c>.
    /// </summary>
    private static readonly HashSet<string> Keywords = ["in", "out", "ref"];

    private const string CtPrefix = "CT_";
    private readonly StringBuilder _sb;
    private readonly Dictionary<string, string> _requiredSimpleTypeTemplate = new();
    private readonly Dictionary<string, string> _optionalSimpleTypeTemplate = new();
    private readonly Dictionary<string, Type> _typeMap = new();
    private int _indentLevel;
    private bool _methodWritten;

    public CodeBuilder(StringBuilder sb)
    {
        _sb = sb;
    }

    internal CodeBuilder AddLine(string s)
    {
        AddIndentedLine(s);
        return this;
    }

    internal CodeBuilder OpenBrace()
    {
        AddIndentedLine("{");
        _indentLevel++;
        return this;
    }

    internal CodeBuilder CloseBrace()
    {
        _indentLevel--;
        AddIndentedLine("}");
        return this;
    }

    internal CodeBuilder StartMethod(string signaturePattern, string typeName)
    {
        if (!typeName.StartsWith(CtPrefix))
            throw new ArgumentException("Type isn't a complex type.", nameof(typeName));

        if (_methodWritten)
            _sb.AppendLine();

        AddIndentation();
        _sb.Append("private ");
        _sb.AppendFormat(signaturePattern, typeName[CtPrefix.Length..]);
        _sb.AppendLine();
        _methodWritten = true;
        return this;
    }

    private void AddIndentedLine(string text)
    {
        AddIndentation();
        _sb.AppendLine(text);
    }

    private void AddIndentation()
    {
        for (var i = 0; i < _indentLevel; i++)
            _sb.Append("    ");
    }

    public override string ToString()
    {
        return _sb.ToString();
    }

    internal string NormalizeCt(string typeName)
    {
        return typeName[3..];
    }

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

    internal string GetSimpleTypeTemplate(string typeName, bool isOptional)
    {
        var templates = isOptional ? _optionalSimpleTypeTemplate : _requiredSimpleTypeTemplate;
        if (!templates.TryGetValue(typeName, out var methodTemplate))
            throw new InvalidOperationException($"Simple type {typeName} ({(isOptional ? "optional" : "required")}) doesn't have defined template.");

        return methodTemplate;
    }

    internal CodeBuilder WriteIndent()
    {
        AddIndentation();
        return this;
    }

    internal CodeBuilder AppendVariable(string variableName)
    {
        _sb.Append(Keywords.Contains(variableName) ? '@' + variableName : variableName);
        return this;
    }

    internal CodeBuilder AppendComplexType(string typeName)
    {
        _sb.Append(typeName[CtPrefix.Length..]);
        return this;
    }

    internal CodeBuilder AppendSimpleType(AttributeElement attribute)
    {
        AppendSimpleType(attribute.Type!);

        var isOptional = attribute.Use is AttributeUseType.Default or AttributeUseType.Optional;
        var nullable = isOptional && attribute.DefaultValue is null;
        var cSharpType = GetSimpleType(attribute.Type!);
        if (nullable && cSharpType.IsValueType)
            _sb.Append('?');

        return this;
    }

    internal CodeBuilder AppendSimpleType(string typeName)
    {
        string? cSharpTypeName = GetSimpleTypeName(typeName);
        _sb.Append(cSharpTypeName);

        return this;
    }

    private string GetSimpleTypeName(string typeName)
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

    private Type GetSimpleType(string typeName)
    {
        return _typeMap[typeName];
    }

    internal CodeBuilder Append(string text)
    {
        _sb.Append(text);
        return this;
    }

    public CodeBuilder EndLine()
    {
        _sb.Append('\n');
        return this;
    }
}
