using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.IO.CodeGen.Model;

namespace ClosedXML.IO.CodeGen;

public class ParserGenerator
{
    private readonly Schema _schema;
    private readonly string _readerName;
    private readonly string _namespaceField;
    private readonly List<string> _parseMethods = new();
    private readonly SchemeTypeMap _typeMap = new();
    private string _targetNamespace = "ClosedXML.Excel.IO";

    public ParserGenerator(Schema schema, string readerField, string nsVariable)
    {
        _schema = schema;
        _readerName = readerField;
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
        _typeMap.AddSimpleTypeTemplate<CSharpType>(typeName, true, methodTemplate);
        return this;
    }

    public ParserGenerator AddSimpleTypeOptional<CSharpType>(string typeName, string methodTemplate)
    {
        _typeMap.AddSimpleTypeTemplate<CSharpType>(typeName, false, methodTemplate);
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
        code.AddLine("using ClosedXML.IO;");
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

    private void GenerateParseMethod(CodeBuilder code, string complexTypeName)
    {
        if (!_schema.TryGetComplexType(complexTypeName, out var complexType))
            throw new InvalidOperationException($"Complex type '{complexTypeName}' not found.");

        complexType.Generate(code, _namespaceField);
    }
}
