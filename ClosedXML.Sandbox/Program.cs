using System;
using System.IO;
using System.Xml;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;

namespace ClosedXML.Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            using var stream = File.Open(@"d:\temp\styles.xml", FileMode.Open, FileAccess.Read);
            using var xmlReader = XmlReader.Create(stream);
            xmlReader.MoveToContent();
            var reader = new XmlTreeReader(xmlReader);

            new StyleSheetReader().Load(reader);
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
