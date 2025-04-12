using NUnit.Framework;
using System.IO;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;

namespace ClosedXML.Tests.IO;

internal class XmlTreeReaderExtensionsTests
{
    private const string AttributeName = "test";

    [TestCase("00000000", 0u)]
    [TestCase("0G000000", null)]
    [TestCase(@"FFFFFFFF", 0xFFFFFFFF)]
    [TestCase(@"FFFFFFFF", 0xFFFFFFFF)]
    [TestCase("abcdef00", 0xABCDEF00)]
    [TestCase("0000000", null)]
    [TestCase(@"", null)]
    [TestCase(@"hello", null)]
    public void GetOptionalUIntHex_parses_8_hex_digits(string xmlText, uint? expectedValue)
    {
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalUIntHex(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

    private static XmlTreeReader CreateReader(string attributeValue)
    {
        var xmlContext = $"<element {AttributeName}=\"{attributeValue}\"/>";
        var xmlReader = XmlReader.Create(new StringReader(xmlContext));
        var reader = new XmlTreeReader(xmlReader, new XmlToEnumMapper.Builder().Build(), true);
        reader.Open("element", string.Empty);
        return reader;
    }
}
