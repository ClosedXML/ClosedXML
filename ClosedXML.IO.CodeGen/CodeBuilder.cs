using ClosedXML.IO.CodeGen.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

    internal string StartComplexTypeParseMethod(string typeName)
    {
        if (!TryGetComplexType(typeName, out var csReturnType))
            csReturnType = "void";

        var signaturePattern = $"{csReturnType} Parse{{0}}(string elementName)";
        StartParseMethod(signaturePattern, NormalizeCt(typeName));
        return csReturnType;
    }

    internal string StartElementGroupParseMethod(string elementGroupName)
    {
        Debug.Assert(elementGroupName.StartsWith(EgPrefix));
        if (!TryGetElementGroup(elementGroupName, out var csReturnType))
            csReturnType = "void";

        var signaturePattern = $"{csReturnType} Parse{{0}}()";
        StartParseMethod(signaturePattern, StripNamePrefix(elementGroupName));
        return csReturnType;
    }

    private void StartParseMethod(string signaturePattern, string methodSuffix)
    {
        AddIndentation();
        _sb.Append("private ");
        _sb.AppendFormat(signaturePattern, methodSuffix);
        _sb.AppendLine();
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

    internal Variable? AppendElementGroupParseCall(string elementGroupName)
    {
        Debug.Assert(elementGroupName.StartsWith(EgPrefix));
        var egName = StripNamePrefix(elementGroupName);
        var parseCall = $"Parse{egName}()";
        WriteIndent();
        if (!TryGetElementGroup(egName, out var csType))
        {
            Append(parseCall).Append(";").EndLine();
            return null;
        }

        var variable = new Variable(csType, egName);
        Append("var ").AppendVariable(variable.Name).Append(" = ").Append(parseCall).Append(";").EndLine();
        return variable;
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

    internal CodeBuilder AppendHookSignature(string referencableName, IReadOnlyList<Variable> parameters)
    {
        WriteIndent().Append("partial void On").Append(StripNamePrefix(referencableName)).Append("Parsed(");

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

    private bool TryGetElementGroup(string elementGroup, [NotNullWhen(true)] out string? csType)
    {
        return _typeMap.TryGetElementGroupCsType(elementGroup, out csType);
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
