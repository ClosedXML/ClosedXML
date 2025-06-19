using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
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

    /// <summary>
    /// Prefix of element groups in XML schema.
    /// </summary>
    private const string EgPrefix = "EG_";

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
        _sb.AppendFormat(signaturePattern, StripNamePrefix(typeName));
        _sb.AppendLine();
        return this;
    }

    internal string NormalizeCt(string typeName)
    {
        if (!typeName.StartsWith(CtPrefix))
            throw new ArgumentException("Type isn't a complex type.", nameof(typeName));

        return typeName[CtPrefix.Length..];
    }

    internal string GetSimpleType(string simpleType)
    {
        return _typeMap.GetSimpleType(simpleType).CsTypeName;
    }

    internal CodeBuilder AppendValue(string simpleType, string value)
    {
        var mappedValue = _typeMap.GetSimpleType(simpleType).MapValue(value);
        _sb.Append(mappedValue);
        return this;
    }

    internal bool TryGetComplexType(string complexType, [NotNullWhen(true)] out string? csType)
    {
        return _typeMap.TryGetComplexTypeCsType(complexType, out csType);
    }

    internal CodeBuilder AppendCallHook(string complexTypeName, IReadOnlyList<Variable> arguments)
    {
        Append("On").AppendComplexType(complexTypeName).Append("Parsed(");
        var isFirst = true;
        foreach (var variable in arguments)
        {
            if (!isFirst)
                Append(", ");

            AppendVariable(variable.Name);
            isFirst = false;
        }

        Append(");").EndLine();
        return this;
    }

    internal CodeBuilder AppendHookSignature(string complexTypeName, IReadOnlyList<Variable> parameters)
    {
        WriteIndent().Append("partial void On").AppendComplexType(complexTypeName).Append("Parsed(");

        var isFirst = true;
        foreach (var parameter in parameters)
        {
            if (!isFirst)
                Append(", ");

            Append(parameter.Type).Append(" ").AppendVariable(parameter.Name);
            isFirst = false;
        }

        Append(");").EndLine();
        return this;
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

    private static string StripNamePrefix(string name)
    {
        if (name.StartsWith(CtPrefix))
            return name[CtPrefix.Length..];

        if (name.StartsWith(EgPrefix))
            return name[EgPrefix.Length..];

        throw new ArgumentException("Name isn't prefixed.");
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
