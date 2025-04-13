using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Intrinsics.X86;
using System.Text;

namespace ClosedXML.IO.CodeGen;

internal class CodeBuilder
{
    private readonly StringBuilder _sb;
    private readonly Dictionary<string, string> _requiredSimpleTypeTemplate = new();
    private readonly Dictionary<string, string> _optionalSimpleTypeTemplate = new();
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

    internal CodeBuilder StartMethod(string signature)
    {
        if (_methodWritten)
            _sb.AppendLine();

        AddIndentation();
        _sb.Append("private ");
        _sb.AppendLine(signature);
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

    internal void AddSimpleTypeTemplate(string typeName, bool isRequired, string methodTemplate)
    {
        if (isRequired)
            _requiredSimpleTypeTemplate.Add(typeName, methodTemplate);
        else
            _optionalSimpleTypeTemplate.Add(typeName, methodTemplate);
    }

    internal string GetSimpleTypeTemplate(string typeName, bool isOptional)
    {
        var templates = isOptional ? _optionalSimpleTypeTemplate : _requiredSimpleTypeTemplate;
        if (!templates.TryGetValue(typeName, out var methodTemplate))
            throw new InvalidOperationException($"Simple type {typeName} ({(isOptional ? "optional" : "required")}) doesn't have defined template.");

        return methodTemplate;
    }
}
