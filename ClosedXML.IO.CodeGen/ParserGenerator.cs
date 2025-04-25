using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.IO.CodeGen.Model;

namespace ClosedXML.IO.CodeGen;

public class ParserGenerator
{
    private readonly string _namespaceField;
    private readonly Schema _schema;
    private readonly string _readerField;
    private readonly List<string> _parseMethods = new();
    private readonly CodeBuilder _code = new(new StringBuilder());
    private string _targetNamespace = "ClosedXML.Excel.IO";

    public ParserGenerator(Schema schema, string readerField, string nsVariable)
    {
        _schema = schema;
        _readerField = readerField;
        _namespaceField = nsVariable;
    }

    public ParserGenerator WithNamespace(string targetNamespace)
    {
        _targetNamespace = targetNamespace;
        return this;
    }

    /// <summary>
    /// Generate <c>Parse*</c> method for a complex type.
    /// </summary>
    /// <param name="complexTypeName">Name of a complex type.</param>
    public ParserGenerator AddParseMethod(string complexTypeName)
    {
        _parseMethods.Add(complexTypeName);
        return this;
    }

    public ParserGenerator AddSimpleTypeRequired<CSharpType>(string typeName, string methodTemplate)
    {
        _code.AddSimpleTypeTemplate<CSharpType>(typeName, true, methodTemplate);
        return this;
    }

    public ParserGenerator AddSimpleTypeOptional<CSharpType>(string typeName, string methodTemplate)
    {
        _code.AddSimpleTypeTemplate<CSharpType>(typeName, false, methodTemplate);
        return this;
    }

    /// <summary>
    /// Generate code from the configuration and a XML schema.
    /// </summary>
    /// <returns>Generated source code.</returns>
    public string Generate()
    {
        _code.AddLine($"namespace {_targetNamespace};");
        _code.EndLine();
        _code.AddLine($"internal partial class {_readerField}");
        _code.OpenBrace();

        foreach (var parseMethod in _parseMethods)
        {
            GenerateParseMethod(parseMethod);
        }

        _code.CloseBrace();
        return _code.ToString();
    }

    private void GenerateParseMethod(string complexTypeName)
    {
        if (!_schema.TryGetComplexType(complexTypeName, out var complexType))
            throw new InvalidOperationException($"Complex type '{complexTypeName}' not found.");

        complexType.GenerateParseMethod(_code, _namespaceField);
    }
}
