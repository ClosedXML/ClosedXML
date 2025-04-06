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
    }
}
