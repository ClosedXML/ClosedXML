using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

public class ReaderConfig
{
    private readonly string _namespaceField = "_ns";
    private readonly Schema _schema;
    private readonly string _readerName;
    private readonly List<string> _parseMethods = new();
    private readonly CodeBuilder _code = new(new StringBuilder());

    public ReaderConfig(Schema schema, string readerName, string nsVariable)
    {
        _schema = schema;
        _readerName = readerName;
        _namespaceField = nsVariable;
    }

    public ReaderConfig AddParseMethod(string complexTypeName)
    {
        _parseMethods.Add(complexTypeName);
        return this;
    }

    public ReaderConfig SkipType(string complexType)
    {
        return this;
    }

    public String Generate()
    {
        _code.AddLine($"public partial class {_readerName}");
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

        switch (complexType)
        {
            case ComplexTypeElement ctElement:
                GenerateParseMethod(ctElement);
                break;
            case ComplexTypeSequence ctSequence:
                GenerateParseMethod(ctSequence);
                break;
            case ComplexTypeChoice ctChoice:
                GenerateParseMethod(ctChoice);
                break;
            default:
                throw new NotImplementedException();
        }
    }

    private void GenerateParseMethod(ComplexTypeChoice complexType)
    {
        var choice = complexType.Choice;
        var min = choice.Occurrences.Min ?? 1;
        var max = choice.Occurrences.Max ?? 1;
        _code.StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        _code.OpenBrace();

        var isFirst = true;
        foreach (var child in choice.Children)
        {
            var element = (ElementType)child;
            var a = isFirst ? string.Empty : "else ";
            isFirst = false;

            _code.AddLine($"{a}if (reader.TryOpen(\"{element.Name}\", {_namespaceField}))");
            _code.OpenBrace();
            _code.AddLine($"Parse{element.TypeName[3..]}(\"{element.Name}\");");
            _code.CloseBrace();
        }

        _code.AddLine("else");
        _code.OpenBrace();
        _code.AddLine("throw PartStructureException.ExpectedChoiceElementNotFound(reader);");
        _code.CloseBrace();
        _code.CloseBrace();
    }

    private void GenerateParseMethod(ComplexTypeSequence complexType)
    {
        var sequence = complexType.Sequence;
        var min = sequence.Occurrences.Min ?? 1;
        var max = sequence.Occurrences.Max ?? 1;
        _code.StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        _code.OpenBrace();
        if (min == 1 && max == 1)
        {
            foreach (var oneOfAttribute in complexType.Attributes)
            {
                if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
                {
                    GenerateReadAttribute(attribute);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }

            foreach (var element in sequence.Children)
            {
                if (element is ElementType elementType)
                {
                    GenerateReadElement(elementType);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }

        }
        else
        {
            throw new NotImplementedException();
        }

        _code.AddLine($"reader.Close(elementName, {_namespaceField});");
        _code.CloseBrace();
    }

    private void GenerateReadElement(ElementType elementType)
    {
        var min = elementType.Occurrences.Min ?? 1;
        var max = elementType.Occurrences.Max ?? 1;

        if (min == 0 && max == int.MaxValue)
        {
            _code.AddLine($"while (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}))");
            _code.OpenBrace();
            _code.AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            _code.CloseBrace();
        }
        else if (min == 1 && max == int.MaxValue)
        {
            _code.AddLine($"reader.Open(\"{elementType.Name}\", {_namespaceField});");
            _code.AddLine("do");
            _code.OpenBrace();
            _code.AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            _code.CloseBrace();
            _code.AddLine($"while (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}));");
        }
        else if (min == 1 && max == 1)
        {
            _code.AddLine($"reader.Open(\"{elementType.Name}\", {_namespaceField}))");
            _code.AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
        }
        else if (min == 0 && max == 1)
        {
            _code.AddLine($"if (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}))");
            _code.OpenBrace();
            _code.AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            _code.CloseBrace();
        }
        else
        {
            throw new NotImplementedException();
        }
    }

    const string prefix = "CT_";

    public void GenerateParseMethod(ComplexTypeElement complexType)
    {
        Debug.Assert(complexType.Name.StartsWith(prefix));
        _code.StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        _code.OpenBrace();
        foreach (var oneOfAttribute in complexType.Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                GenerateReadAttribute(attribute);
            }
            else
            {
                throw new NotImplementedException();
            }
        }
        _code.AddLine($"reader.Close(elementName, {_namespaceField});");
        _code.CloseBrace();
    }

    private void GenerateReadAttribute(AttributeElement attribute)
    {
        Debug.Assert(attribute.Name is not null);
        Debug.Assert(attribute.Type is not null);
        var isOptional = attribute.Use != AttributeUseType.Required;

        if (isOptional)
        {
            if (_optionalSimpleTypeTemplate.TryGetValue(attribute.Type, out var methodTemplate))
            {
                var b = "var " + EscapeVariableName(attribute.Name) + " = " + string.Format(methodTemplate, attribute.Name, attribute.DefaultValue ?? "null") + ";";
                _code.AddLine(b);
                return;
            }
            throw new NotImplementedException($"Optional {attribute.Type}");
        }
        else
        {
            if (_requiredSimpleTypeTemplate.TryGetValue(attribute.Type, out var methodTemplate))
            {
                var b = "var " + EscapeVariableName(attribute.Name) + " = " + string.Format(methodTemplate, attribute.Name) + ";";
                _code.AddLine(b);
                return;
            }
            throw new NotImplementedException($"Required {attribute.Type}");
        }
    }

    private readonly Dictionary<string, string> _requiredSimpleTypeTemplate = new();
    private readonly Dictionary<string, string> _optionalSimpleTypeTemplate = new();

    public ReaderConfig AddSimpleTypeRequired(string typeName, string methodTemplate)
    {
        _requiredSimpleTypeTemplate.Add(typeName, methodTemplate);
        return this;
    }

    public ReaderConfig AddSimpleTypeOptional(string typeName, string methodTemplate)
    {
        _optionalSimpleTypeTemplate.Add(typeName, methodTemplate);
        return this;
    }

    private string EscapeVariableName(string name)
    {
        if (name == "in") return "@in";
        return name;
    }
}
