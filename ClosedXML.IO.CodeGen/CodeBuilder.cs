using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.IO.CodeGen;

internal class CodeBuilder
{
    /// <summary>
    /// C# keywords. The variables with that name must be escaped, e.g. <c>in</c> must be
    /// <c>@in</c>.
    /// </summary>
    private static readonly HashSet<string> Keywords = ["in", "out", "ref"];

    /// <summary>
    /// Prefix of complex types in XML schema.
    /// </summary>
    private const string CtPrefix = "CT_";

    private readonly SchemeTypeMap _typeMap;
    private readonly StringBuilder _sb;
    private int _indentLevel;

    public CodeBuilder(StringBuilder sb, SchemeTypeMap typeMap)
    {
        _sb = sb;
        _typeMap = typeMap;
    }

    internal CodeBuilder AddLine(string lineText)
    {
        AddIndentedLine(lineText);
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

    internal CodeBuilder Append(string text)
    {
        _sb.Append(text);
        return this;
    }

    internal CodeBuilder EndLine()
    {
        _sb.AppendLine();
        return this;
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

    internal CodeBuilder StartMethod(string signaturePattern, string typeName)
    {
        AddIndentation();
        _sb.Append("private ");
        _sb.AppendFormat(signaturePattern, NormalizeCt(typeName));
        _sb.AppendLine();
        return this;
    }

    internal string NormalizeCt(string typeName)
    {
        if (!typeName.StartsWith(CtPrefix))
            throw new ArgumentException("Type isn't a complex type.", nameof(typeName));

        return typeName[CtPrefix.Length..];
    }

    internal CodeBuilder AppendComplexType(string typeName)
    {
        _sb.Append(NormalizeCt(typeName));
        return this;
    }

    internal CodeBuilder AppendSimpleTypeMethod(AttributeElement attribute)
    {
        var codeFragment = _typeMap.GetSimpleTypeMethod(attribute);
        return Append(codeFragment);
    }

    internal CodeBuilder AppendSimpleType(AttributeElement attribute)
    {
        AppendSimpleType(attribute.Type!);

        var isOptional = attribute.Use is AttributeUseType.Default or AttributeUseType.Optional;
        var nullable = isOptional && attribute.DefaultValue is null;
        if (nullable)
            _sb.Append('?');

        return this;
    }

    private void AppendSimpleType(string typeName)
    {
        var cSharpTypeName = _typeMap.GetSimpleTypeName(typeName);
        _sb.Append(cSharpTypeName);
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
}
