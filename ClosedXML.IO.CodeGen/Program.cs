using System;
using System.Diagnostics;
using System.IO;
using System.Text;
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
        using var reader = new XmlTreeReader(fileStream, new XsdEnumMapper(), false);
        var parser = new XsdSchemaParser();

        var schema = parser.ParseSchema(reader);

        Console.Out.WriteLine($"File {args[0]} successfully parsed.");

        var sb = new StringBuilder();
        var visitor = new XsdCopyVisitor(sb);
        visitor.Visit(schema);

        File.WriteAllText(args[1], sb.ToString());

        Console.WriteLine($"Wrote copy to {args[1]}");

        var cacheRecordsGenerator = new ParserGenerator(schema, "PivotCacheRecordsReader", "_ns")
            .WithNamespace("ClosedXML.Excel.IO")
            .AddSimpleTypeRequired<uint>("xsd:unsignedInt", "_reader.GetUInt(\"{0}\")")
            .AddSimpleTypeOptional<int?>("xsd:int", "_reader.GetOptionalInt(\"{0}\")")
            .AddSimpleTypeRequired<bool>("xsd:boolean", "_reader.GetBool(\"{0}\")")
            .AddSimpleTypeOptional<bool?>("xsd:boolean", "_reader.GetOptionalBool(\"{0}\")")
            .AddSimpleTypeOptional<string?>("s:ST_Xstring", "_reader.GetOptionalXString(\"{0}\")")
            .AddSimpleTypeRequired<string>("s:ST_Xstring", "_reader.GetXString(\"{0}\")")
            .AddSimpleTypeOptional<uint?>("xsd:unsignedInt", "_reader.GetOptionalUInt(\"{0}\")")
            .AddSimpleTypeRequired<DateTime>("xsd:dateTime", "_reader.GetDateTime(\"{0}\")")
            .AddSimpleTypeOptional<uint?>("ST_UnsignedIntHex", "_reader.GetOptionalUIntHex(\"{0}\")")
            .AddSimpleTypeRequired<double>("xsd:double", "_reader.GetDouble(\"{0}\")")

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
        Console.ReadKey();
    }
}
