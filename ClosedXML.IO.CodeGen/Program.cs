using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.TopLevel;
using ClosedXML.IO.CodeGen.XsdParser;

namespace ClosedXML.IO.CodeGen;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine($"    {Process.GetCurrentProcess().ProcessName}.exe name-of-ooxml.xsd output-path.xsd");
            Console.Error.WriteLine();
            return;
        }

        using var fileStream = File.OpenRead(args[0]);
        using var xmlReader = XmlReader.Create(fileStream);
        using var reader = new XmlTreeReader(xmlReader, new XsdEnumMapper());
        var parser = new XsdSchemaParser();

        var schema = parser.ParseSchema(reader);

        Console.Out.WriteLine($"File {args[0]} successfully parsed.");

        var sb = new StringBuilder();
        var visitor = new XsdCopyVisitor(sb);
        visitor.Visit(schema);

        File.WriteAllText(args[1], sb.ToString());

        Console.WriteLine($"Wrote copy to {args[1]}");

        var cacheRecords = new ReaderConfig(schema, "PivotCacheRecordsReader", "_ns")
            .AddSimpleTypeRequired("xsd:unsignedInt", "reader.GetBool(\"{0}\")")
            .AddSimpleTypeOptional("xsd:int", "reader.GetOptionalInt(\"{0}\") ?? {1}")
            .AddSimpleTypeRequired("xsd:boolean", "reader.GetBool(\"{0}\")")
            .AddSimpleTypeOptional("xsd:boolean", "reader.GetOptionalBool(\"{0}\") ?? {1}")
            .AddSimpleTypeOptional("s:ST_Xstring", "reader.GetOptionalXString(\"{0}\") ?? {1}")
            .AddSimpleTypeRequired("s:ST_Xstring", "reader.GetXString(\"{0}\")")
            .AddSimpleTypeOptional("xsd:unsignedInt", "reader.GetOptionalUint(\"{0}\") ?? {1}")
            .AddSimpleTypeRequired("xsd:dateTime", "reader.GetDateTime(\"{0}\")")
            .AddSimpleTypeOptional("ST_UnsignedIntHex", "reader.GetOptionalUintHex(\"{0}\") ?? {1}")
            .AddSimpleTypeRequired("xsd:double", "reader.GetDouble(\"{0}\")")

            .AddParseMethod("CT_PivotCacheRecords")
            .AddParseMethod("CT_Record")
            .AddParseMethod("CT_Missing")
            .AddParseMethod("CT_Number")
            .AddParseMethod("CT_Boolean")
            .AddParseMethod("CT_Error")
            .AddParseMethod("CT_String")
            .AddParseMethod("CT_DateTime")
            .AddParseMethod("CT_Index")
            .AddParseMethod("CT_X")
            .AddParseMethod("CT_Tuples")
            .AddParseMethod("CT_Tuple")
            ;

        var cacheRecordsSource = cacheRecords.Generate();
        Console.WriteLine(cacheRecordsSource);
        Console.ReadKey();
    }
}

public class ReaderConfig
{
    private readonly string _namespaceField = "_ns";
    private readonly Schema _schema;
    private readonly string _readerName;
    private readonly List<string> _parseMethods = new();
    private readonly StringBuilder _sb = new();

    private int _indentLevel = 0;
    private bool _methodWritten = false;

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

    public StringBuilder Generate()
    {
        AddLine($"public partial class {_readerName}");
        OpenBrace();

        foreach (var parseMethod in _parseMethods)
        {
            GenerateParseMethod(parseMethod);
        }

        CloseBrace();
        return _sb;
    }

    private void AddLine(string s)
    {
        AppendIndent();
        _sb.AppendLine(s);
    }

    private void OpenBrace()
    {
        AppendIndent();
        _sb.AppendLine("{");
        _indentLevel++;
    }

    private void CloseBrace()
    {
        _indentLevel--;
        AppendIndent();
        _sb.AppendLine("}");
    }

    private void AppendIndent()
    {
        for (var i = 0; i < _indentLevel; i++)
            _sb.Append("    ");
    }
    private void StartMethod(string s)
    {
        if (_methodWritten)
            _sb.AppendLine();

        AddLine(s);
        _methodWritten = true;
    }


    private void GenerateParseMethod(string complexTypeName)
    {
        var complexType = _schema.Entries.OfType<ComplexType>().SingleOrDefault(x => x.Name == complexTypeName);
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
        StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        OpenBrace();

        var isFirst = true;
        foreach (var child in choice.Children)
        {
            var element = (ElementType)child;
            var a = isFirst ? string.Empty : "else ";
            isFirst = false;

            AddLine($"{a}if (reader.TryOpen(\"{element.Name}\", {_namespaceField}))");
            OpenBrace();
            AddLine($"Parse{element.TypeName[3..]}(\"{element.Name}\");");
            CloseBrace();
        }

        AddLine("else");
        OpenBrace();
        AddLine("throw PartStructureException.ExpectedChoiceElementNotFound(reader);");
        CloseBrace();

        CloseBrace();
    }

    private void GenerateParseMethod(ComplexTypeSequence complexType)
    {
        var sequence = complexType.Sequence;
        var min = sequence.Occurrences.Min ?? 1;
        var max = sequence.Occurrences.Max ?? 1;
        StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        OpenBrace();
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

        AddLine($"reader.Close(elementName, {_namespaceField});");
        CloseBrace();
    }

    private void GenerateReadElement(ElementType elementType)
    {
        var min = elementType.Occurrences.Min ?? 1;
        var max = elementType.Occurrences.Max ?? 1;

        if (min == 0 && max == int.MaxValue)
        {
            AddLine($"while (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}))");
            OpenBrace();
            AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            CloseBrace();
        }
        else if (min == 1 && max == int.MaxValue)
        {
            AddLine($"reader.Open(\"{elementType.Name}\", {_namespaceField});");
            AddLine("do");
            OpenBrace();
            AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            CloseBrace();
            AddLine($"while (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}));");
        }
        else if (min == 1 && max == 1)
        {
            AddLine($"reader.Open(\"{elementType.Name}\", {_namespaceField}))");
            AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
        }
        else if (min == 0 && max == 1)
        {
            AddLine($"if (reader.TryOpen(\"{elementType.Name}\", {_namespaceField}))");
            OpenBrace();
            AddLine($"Parse{elementType.TypeName[3..]}(\"{elementType.Name}\");");
            CloseBrace();
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
        StartMethod($"void Parse{complexType.Name[prefix.Length..]}(string elementName)");
        OpenBrace();
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
        AddLine($"reader.Close(elementName, {_namespaceField});");
        CloseBrace();
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
                AddLine(b);
                return;
            }
            throw new NotImplementedException($"Optional {attribute.Type}");
        }
        else
        {
            if (_requiredSimpleTypeTemplate.TryGetValue(attribute.Type, out var methodTemplate))
            {
                var b = "var " + EscapeVariableName(attribute.Name) + " = " + string.Format(methodTemplate, attribute.Name) + ";";
                AddLine(b);
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
