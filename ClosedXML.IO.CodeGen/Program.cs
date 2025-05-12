using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.XsdParser;

namespace ClosedXML.IO.CodeGen;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine($"    {Process.GetCurrentProcess().ProcessName}.exe name-of-ooxml.xsd command");
            Console.Error.WriteLine();
            return;
        }

        var schemaPath = args[0];
        using var fileStream = File.OpenRead(schemaPath);
        using var reader = new XmlTreeReader(fileStream, new XsdEnumMapper(), false);
        var parser = new XsdSchemaParser();

        var schema = parser.ParseSchema(reader);

        Console.Out.WriteLine($"Schema {schemaPath} successfully parsed.");

        var command = args[1];
        switch (command)
        {
            case "copy":
                var sb = new StringBuilder();
                var visitor = new XsdCopyVisitor(sb);
                visitor.Visit(schema);
                var outputFile = Path.ChangeExtension(schemaPath, "copy");
                File.WriteAllText(outputFile, sb.ToString());
                Console.WriteLine($"Wrote copy to {outputFile}");
                break;

            case "styles":
                GenerateStylesReader(schema);
                break;

            case "cache-records":
                GenerateCacheRecords(schema);
                break;

            default:
                Console.WriteLine($"Unknown command '{command}'");
                break;
        }

        Console.ReadKey();
    }

    private static void GenerateStylesReader(Schema schema)
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
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_FontId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
            })
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
            .AddSimpleTypeEnum("ST_TableStyleType", "XLTableStyleType")
            .AddComplexTypeMapping("CT_Color", "XLColor")
            .AddComplexTypeMapping("CT_GradientStop", "(FractionOfOne Value, XLColor Color)")
            .AddComplexTypeMapping("CT_BorderPr", "XLBorderLine")
            .AddComplexTypeMapping("CT_PatternFill", "XLFillFormat")
            .AddComplexTypeMapping("CT_GradientFill", "XLFillFormat")
            .AddComplexTypeMapping("CT_NumFmt", "(int NumFmtId, string FormatCode)")
            .AddComplexTypeMapping("CT_CellAlignment", "XLAlignmentFormat")
            .AddComplexTypeMapping("CT_CellProtection", "XLProtectionFormat")
            ;

        var stylesReaderGenerator = new ParserGenerator(schema, typeMap, "StylesReader", "_ns")
            .AddUsing("System.Collections.Generic")
            .AddUsing("ClosedXML.IO")
            .AddUsing("ClosedXML.Excel.Formatting")
            .AddParseMethod("CT_Stylesheet")
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
        Console.WriteLine(stylesReaderSource);
    }

    private static void GenerateCacheRecords(Schema schema)
    {
        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes();

        var cacheRecordsGenerator = new ParserGenerator(schema, typeMap, "PivotCacheRecordsReader", "_ns")
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
        Console.WriteLine(cacheRecordsSource);
    }
}
