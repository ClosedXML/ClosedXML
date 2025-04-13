using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml;
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

        var cacheRecords = new ParserGenerator(schema, "PivotCacheRecordsReader", "_ns")
            .AddSimpleTypeRequired("xsd:unsignedInt", "reader.GetUInt(\"{0}\")")
            .AddSimpleTypeOptional("xsd:int", "reader.GetOptionalInt(\"{0}\")")
            .AddSimpleTypeRequired("xsd:boolean", "reader.GetBool(\"{0}\")")
            .AddSimpleTypeOptional("xsd:boolean", "reader.GetOptionalBool(\"{0}\")")
            .AddSimpleTypeOptional("s:ST_Xstring", "reader.GetOptionalXString(\"{0}\")")
            .AddSimpleTypeRequired("s:ST_Xstring", "reader.GetXString(\"{0}\")")
            .AddSimpleTypeOptional("xsd:unsignedInt", "reader.GetOptionalUInt(\"{0}\")")
            .AddSimpleTypeRequired("xsd:dateTime", "reader.GetDateTime(\"{0}\")")
            .AddSimpleTypeOptional("ST_UnsignedIntHex", "reader.GetOptionalUIntHex(\"{0}\")")
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
