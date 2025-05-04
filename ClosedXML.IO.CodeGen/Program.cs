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
            .AddSimpleTypeRequired("ST_NumFmtId", "_reader.GetUInt(\"{0}\")", "uint")
            .AddSimpleTypeOptional("ST_PatternType", "_reader.GetOptionalEnum<XLFillPatternValues>(\"{0}\")", "XLFillPatternValues")
            ;

        var stylesReaderGenerator = new ParserGenerator(schema, typeMap, "StylesPartReader", "_ns")
            .AddParseMethod("CT_PatternFill")
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
