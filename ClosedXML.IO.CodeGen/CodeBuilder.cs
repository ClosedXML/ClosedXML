using ClosedXML.IO.CodeGen.Model;
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

    internal string StartParseMethod(ParsletName name, params string[] parameters)
    {
        if (!TryGetCsType(name, out var csReturnType))
            csReturnType = "void";

        AddIndentedLine($"private {csReturnType} Parse{name.WithoutPrefix()}({string.Join(", ", parameters)})");
        return csReturnType;
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

    internal bool TryGetCsType(ParsletName name, [NotNullWhen(true)] out string? csType)
    {
        return _typeMap.TryGetParsletCsType(name, out csType);
    }

    internal Variable? AddParseCall(ParsletName name, string variableName, string[] arguments)
    {
        if (!TryGetCsType(name, out var csType))
        {
            WriteIndent().AppendParseCall(name, arguments).Append(";").EndLine();
            return null;
        }

        WriteIndent().Append("var ").AppendVariable(variableName).Append(" = ").AppendParseCall(name, arguments).Append(";").EndLine();
        return new Variable(csType, variableName);
    }

    internal CodeBuilder AppendParseCall(ParsletName name, string[] arguments)
    {
        return Append($"Parse{name.WithoutPrefix()}({string.Join(", ", arguments)})");
    }

    internal CodeBuilder AppendCallHook(ParsletName name, IReadOnlyList<Variable> arguments)
    {
        Append("On").Append(name.WithoutPrefix()).Append("Parsed(");
        var isFirst = true;
        foreach (var variable in arguments)
        {
            if (!isFirst)
                Append(", ");

            AppendVariable(variable.Name);
            isFirst = false;
        }

        Append(")");
        return this;
    }

    internal CodeBuilder AppendHookSignature(ParsletName name, IReadOnlyList<Variable> parameters)
    {
        WriteIndent().Append("partial void On").Append(name.WithoutPrefix()).Append("Parsed(");

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

    internal CodeBuilder AppendSimpleTypeMethod(AttributeElement attribute)
    {
        var codeFragment = _typeMap.GetSimpleTypeMethod(attribute);
        return Append(codeFragment);
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
