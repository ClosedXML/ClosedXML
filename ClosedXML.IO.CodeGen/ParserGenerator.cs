using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.IO.CodeGen.Model;

namespace ClosedXML.IO.CodeGen;

internal class ParserGenerator
{
    private readonly Schema _schema;
    private readonly string _readerName;
    private readonly string _namespaceField;
    private readonly List<ParsletName> _parseMethods = new();
    private readonly SchemeTypeMap _typeMap;
    private readonly List<string> _usings = new();
    private string _targetNamespace = "ClosedXML.Excel.IO";

    internal ParserGenerator(Schema schema, SchemeTypeMap typeMap, string readerField, string nsVariable)
    {
        _schema = schema;
        _typeMap = typeMap;
        _readerName = readerField;
        _namespaceField = nsVariable;
    }

    public ParserGenerator WithNamespace(string targetNamespace)
    {
        _targetNamespace = targetNamespace;
        return this;
    }

    public ParserGenerator AddUsing(string usingNamespace)
    {
        _usings.Add(usingNamespace);
        return this;
    }

    /// <summary>
    /// Generate <c>Parse*</c> method for a top-level element in the XSD file.
    /// </summary>
    /// <param name="name">Name of a complex type or element group.</param>
    public ParserGenerator AddParseMethod(ParsletName name)
    {
        _parseMethods.Add(name);
        return this;
    }

    /// <summary>
    /// Generate code from the configuration and a XML schema.
    /// </summary>
    /// <returns>Generated source code.</returns>
    public string Generate()
    {
        var code = new CodeBuilder(new StringBuilder(), _typeMap);
        code.AddLine("#nullable enable");
        code.EndLine();
        foreach (var usingNs in _usings)
            code.AddLine($"using {usingNs};");

        code.EndLine();
        code.AddLine($"namespace {_targetNamespace};");
        code.EndLine();
        code.AddLine($"internal partial class {_readerName}");
        code.OpenBrace();

        if (_parseMethods.Count > 0)
            GenerateParseMethod(code, _parseMethods[0]);

        foreach (var parseMethod in _parseMethods[1..])
        {
            code.EndLine();
            GenerateParseMethod(code, parseMethod);
        }

        code.CloseBrace();
        return code.ToString();
    }

    private void GenerateParseMethod(CodeBuilder code, ParsletName parsletName)
    {
        if (!_schema.TryGetParslet(parsletName, out var parslet))
            throw new InvalidOperationException($"Unable to find definition for '{parsletName.Value}'. Was it part of the XSD file?");

        parslet.GenerateParseMethod(code, _namespaceField);
    }
}
