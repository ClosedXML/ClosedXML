using NUnit.Framework;
using System.IO;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;

namespace ClosedXML.Tests.IO;

internal class XmlTreeReaderExtensionsTests
{
    private const string AttributeName = "test";

    [Test]
    public void GetXString_throws_when_attribute_is_not_present()
    {
        using var reader = CreateReader("dummy");

        var ex = Assert.Throws<PartStructureException>(() => reader.GetXString("nonexistent"));
        Assert.That(ex, Has.Message.Contain("XML doesn't contain a required attribute 'nonexistent'."));
    }


    [TestCase("&amp;", "&")]
    [TestCase("_x0009_", "\t")]
    [TestCase("_X0009_", "_X0009_")]
    [TestCase("Hello &lt;user&gt; - _x0045__x004F__x004C_", "Hello <user> - EOL")]
    public void GetOptionalXString_returns_XString_decoded_xml_decoded_text(string xmlText, string expectedValue)
    {
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalXString(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

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
