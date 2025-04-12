using System.IO;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;

namespace ClosedXML.Tests.IO;

/// <summary>
/// Test various methods (including extension methods) that reader correctly reads the value of
/// an attribute.
/// </summary>
internal class XmlTreeReaderAttributesTests
{
    private const string AttributeName = "test";

    [TestCase("true", true)]
    [TestCase("1", true)]
    [TestCase("false", false)]
    [TestCase("0", false)]
    [TestCase("some text", null)]
    [TestCase("TRUE", null)] // xsd says case sensitive, for non-readable values return null
    [TestCase("FALSE", null)]
    public void GetOptionalBool_reads_xsd_compliant_bool_values(string xmlText, bool? expectedValue)
    {
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalBool(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

    [TestCase("0", 0)]
    [TestCase("17", 17)]
    [TestCase("2147483647", 2147483647)]
    [TestCase("-2147483648", -2147483648)]
    [TestCase("+7", 7)] // Canonical representation forbids plus sign or leading zeros, but they are readable
    [TestCase("05", 5)]
    [TestCase("", null)]
    [TestCase("3.0", null)]
    [TestCase("2147483648", null)]
    [TestCase("-2147483649", null)]
    [TestCase("one", null)]
    public void GetOptionalInt_reads_xsd_compliant_int_values(string xmlText, int? expectedValue)
    {
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalInt(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

    private static XmlTreeReader CreateReader(string attributeValue, XmlToEnumMapper mapper = null)
    {
        var xmlContext = $"<element {AttributeName}=\"{attributeValue}\"/>";
        var xmlReader = XmlReader.Create(new StringReader(xmlContext));
        mapper ??= new XmlToEnumMapper.Builder().Build();
        var reader = new XmlTreeReader(xmlReader, mapper, true);
        reader.Open("element", string.Empty);
        return reader;
    }
}
