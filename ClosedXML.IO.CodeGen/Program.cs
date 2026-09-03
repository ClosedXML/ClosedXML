using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.TopLevel;
using ClosedXML.IO.CodeGen.XsdParser;

namespace ClosedXML.IO.CodeGen;

public static class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine($"    {Process.GetCurrentProcess().ProcessName}.exe command name-of-ooxml.xsd output.cs");
            Console.Error.WriteLine();
            return;
        }

        var command = args[0];
        var schemaPath = args[1];
        var target = args[2];
        using var fileStream = File.OpenRead(schemaPath);
        using var reader = new XmlTreeReader(fileStream, new XsdEnumMapper(), true);
        var parser = new XsdSchemaParser();

        var schema = parser.ParseSchema(reader);
        switch (command)
        {
            case "styles":
                GenerateStylesReader(schema, target);
                break;

            case "cache-records":
                GenerateCacheRecords(schema, target);
                break;

            case "partial-rst":
                GeneratePartialRstReader(schema, target);
                break;

            default:
                Console.WriteLine($"Unknown command '{command}'");
                break;
        }

        Console.ReadKey();
    }

    private static void GenerateStylesReader(Schema schema, string target)
    {
        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_NumFmtId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddStFontId()
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_FillId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_BorderId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_CellStyleXfId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_TextRotation",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_DxfId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
            .AddSimpleTypeEnum("ST_PatternType", "XLFillPatternValues")
            .AddSimpleTypeEnum("ST_GradientType", "XLGradientType", "linear", "XLGradientType.Linear")
            .AddSimpleTypeEnum("ST_BorderStyle", "XLBorderStyleValues", "none", "XLBorderStyleValues.None")
            .AddSimpleTypeEnum("ST_HorizontalAlignment", "XLAlignmentHorizontalValues")
            .AddSimpleTypeEnum("ST_VerticalAlignment", "XLAlignmentVerticalValues", "bottom", "XLAlignmentVerticalValues.Bottom")
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_TableStyleType",
                CsTypeName = "(XLTableStyleRegionValues?, XLPivotStyleRegionValues?)",
                RequiredTemplate = "_reader.GetStringMappedValue(\"{0}\", TableStyleTypeMap)"
            })
            .AddComplexTypeMapping("CT_Color", "XLColor")
            .AddComplexTypeMapping("CT_GradientStop", "(FractionOfOne Value, XLColor Color)")
            .AddComplexTypeMapping("CT_Font", "XLDifferentialFontValue")
            .AddComplexTypeMapping("CT_Fill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_Border", "XLDifferentialBorderValue")
            .AddComplexTypeMapping("CT_BorderPr", "XLBorderLine")
            .AddComplexTypeMapping("CT_PatternFill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_GradientFill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_NumFmt", "(int NumFmtId, XLNumberFormat Format)")
            .AddComplexTypeMapping("CT_CellAlignment", "XLDifferentialAlignmentValue")
            .AddComplexTypeMapping("CT_CellProtection", "XLDifferentialProtectionValue")
            .AddComplexTypeMapping("CT_Xf", "(XLCellFormatValue Format, int? CellStyleXfId)")
            .AddComplexTypeMapping("CT_CellXfs", "List<(XLCellFormatValue Format, int? CellStyleXfId)>")
            .AddComplexTypeMapping("CT_CellStyle", "(int CellStyleXfId, XLCellStyleValue Style)")
            .AddComplexTypeMapping("CT_CellStyles", "Dictionary<int, XLCellStyleValue>")
            .AddComplexTypeMapping("CT_RgbColor", "uint")
            ;

        var stylesReaderGenerator = new ParserGenerator(schema, typeMap, "StylesReader")
            .AddUsing("System.Collections.Generic")
            .AddUsing("ClosedXML.IO")
            .AddUsing("ClosedXML.Excel.Formatting")
            //.AddParseMethod("CT_Stylesheet")
            .AddParseMethod("CT_NumFmts")
            .AddParseMethod("CT_NumFmt")
            .AddParseMethod("CT_Fonts")
            // AddParseMethod("CT_Font")
            .AddParseMethod("CT_Fills")
            .AddParseMethod("CT_Fill")
            .AddParseMethod("CT_PatternFill")
            .AddParseMethod("CT_GradientFill")
            .AddParseMethod("CT_GradientStop")
            .AddParseMethod("CT_Borders")
            .AddParseMethod("CT_Border")
            .AddParseMethod("CT_BorderPr")
            .AddParseMethod("CT_CellStyleXfs")
            .AddParseMethod("CT_Xf")
            .AddParseMethod("CT_CellAlignment")
            .AddParseMethod("CT_CellProtection")
            .AddParseMethod("CT_CellXfs")
            .AddParseMethod("CT_CellStyles")
            .AddParseMethod("CT_CellStyle")
            .AddParseMethod("CT_Dxfs")
            .AddParseMethod("CT_Dxf")
            .AddParseMethod("CT_TableStyles")
            .AddParseMethod("CT_TableStyle")
            .AddParseMethod("CT_TableStyleElement")
            .AddParseMethod("CT_Colors")
            .AddParseMethod("CT_IndexedColors")
            .AddParseMethod("CT_MRUColors")
            .AddParseMethod("CT_RgbColor")
            ;

        var stylesReaderSource = stylesReaderGenerator.Generate();
        File.WriteAllText(target, stylesReaderSource);
        Console.WriteLine(stylesReaderSource);
    }

    private static void GenerateCacheRecords(Schema schema, string target)
    {
        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes();

        var cacheRecordsGenerator = new ParserGenerator(schema, typeMap, "PivotCacheRecordsReader")
            .WithNamespace("ClosedXML.Excel.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("ClosedXML.IO")

            // CT_PivotCacheRecords - hand-coded
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

        var cacheRecordsSource = cacheRecordsGenerator.Generate();
        File.WriteAllText(target, cacheRecordsSource);
        Console.WriteLine(cacheRecordsSource);
    }

    /// <summary>
    /// Generate a reader to read <c>CT_Rst</c>. It is a partial reader used in other readers.
    /// </summary>
    private static void GeneratePartialRstReader(Schema schema, string target)
    {
        NormalizeChoicesZeroToOne(schema, "CT_RPrElt");

        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddStFontId()
            .AddSimpleTypeEnum("ST_PhoneticType", "XLPhoneticType", "fullwidthKatakana", "XLPhoneticType.FullWidthKatakana")
            .AddSimpleTypeEnum("ST_PhoneticAlignment", "XLPhoneticAlignment", "left", "XLPhoneticAlignment.Left")
            .AddSimpleTypeEnum("s:ST_VerticalAlignRun", "XLFontVerticalTextAlignmentValues")
            .AddSimpleTypeEnum("ST_FontScheme", "XLFontScheme")
            .AddSimpleTypeEnum("ST_UnderlineValues", "XLFontUnderlineValues", "single", "XLFontUnderlineValues.Single")

            .AddComplexTypeMapping("CT_Color", "XLColor")
            .AddComplexTypeMapping("CT_BooleanProperty", "bool")
            .AddComplexTypeMapping("CT_IntProperty", "int")
            .AddComplexTypeMapping("CT_FontSize", "XLFontSize")
            .AddComplexTypeMapping("CT_UnderlineProperty", "XLFontUnderlineValues")
            .AddComplexTypeMapping("CT_VerticalAlignFontProperty", "XLFontVerticalTextAlignmentValues")
            .AddComplexTypeMapping("CT_FontScheme", "XLFontScheme")
            .AddComplexTypeMapping("CT_FontName", "XLFontName")

            .AddComplexTypeMapping("s:ST_Xstring", "string")
            .AddComplexTypeMapping("CT_RPrElt", "XLDifferentialFontValue", "Unit")
            .AddComplexTypeMapping("CT_RElt", "(string Text, XLDifferentialFontValue Font)")
            .AddComplexTypeMapping("CT_PhoneticRun", "PhoneticRunDto")
            .AddComplexTypeMapping("CT_PhoneticPr", "PhoneticProperties")
            .AddComplexTypeMapping("CT_Rst", "OneOf<string, XLImmutableRichText>")

            .AddParseCall("s:ST_Xstring", "_reader.ParseXString")
            ;

        var rstGenerator = new ParserGenerator(schema, typeMap, "RstReader")
            .WithNamespace("ClosedXML.Excel.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("ClosedXML.Excel.Formatting")
            .AddUsing("ClosedXML.Excel.CalcEngine")
            .AddUsing("ClosedXML.IO")
            .AddUsing("PhoneticProperties = ClosedXML.Excel.XLImmutableRichText.PhoneticProperties")
            .AddUsing("PhoneticRunDto = (string Text, int StartIndex, int EndIndex)")

            .AddParseMethod("CT_PhoneticRun") // 1816
            .AddParseMethod("CT_RElt") // 1823
            .AddParseMethod("CT_RPrElt") // 1829
            .AddParseMethod("CT_Rst") // 1849
            .AddParseMethod("CT_PhoneticPr") // 1857
            .AddParseMethod("CT_BooleanProperty") // 3751
            .AddParseMethod("CT_FontSize") // 3754
            .AddParseMethod("CT_IntProperty") // 3757
            .AddParseMethod("CT_FontName") // 3760
            .AddParseMethod("CT_VerticalAlignFontProperty") // 3763
            .AddParseMethod("CT_FontScheme") // 3766
            .AddParseMethod("CT_UnderlineProperty") // 3776
            ;

        var rstSource = rstGenerator.Generate();
        File.WriteAllText(target, rstSource);
        Console.WriteLine(rstSource);
    }

    /// <summary>
    /// Normalize a choice that has multiple elements with 0..1 cardinality into the 1..1
    /// cardinality. Semantically, it's identical.
    /// </summary>
    private static void NormalizeChoicesZeroToOne(Schema schema, ParsletName name)
    {
        foreach (var choice in schema.Entries.OfType<ComplexTypeChoice>().Where(x => name == x.Name))
        {
            var choiceChildren = choice.Choice.Children;
            var exhibitsProblem = choiceChildren.All(x => x is ElementType
            {
                Occurrences: { ActualMin: 0, ActualMax: 1 }
            });
            if (!exhibitsProblem)
            {
                throw new InvalidOperationException($"Type {name} does not exhibit expected problem.");
            }

            var fixedChildren = choiceChildren.Cast<ElementType>().Select(x => new ElementType
            {
                Name = x.Name,
                Occurrences = new Occurrences(1, 1),
                TypeName = x.TypeName
            }).ToList();

            // Replace the original children with new ones
            choice.Choice.Children.Clear();
            choice.Choice.Children.AddRange(fixedChildren);
        }
    }

    public static SchemeTypeMap AddStFontId(this SchemeTypeMap schemeTypeMap)
    {
        return schemeTypeMap.AddSimpleType(new SimpleTypeMapping
        {
            Name = "ST_FontId",
            CsTypeName = "uint",
            RequiredTemplate = "_reader.GetUInt(\"{0}\")",
            OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
        });
    }
}
