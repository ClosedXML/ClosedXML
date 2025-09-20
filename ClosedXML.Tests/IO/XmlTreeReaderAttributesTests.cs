using System;
using System.IO;
using System.Reflection;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using ClosedXML.Utils;
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

    [TestCase("0", 0u)]
    [TestCase("57", 57u)]
    [TestCase("2147483647", 2147483647u)]
    [TestCase("4294967295", 4294967295u)]
    [TestCase("-7", null)]
    [TestCase("value", null)]
    [TestCase("4294967296", null)] // One above max value
    [TestCase("9223372036854775808", null)]
    public void GetOptionalUint_reads_xsd_compliant_unsignedInt_values(string xmlText, uint? expectedValue)
    {
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalUInt(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

    [TestCase("0", 0)]
    [TestCase("1.75", 1.75)]
    [TestCase("-1.75e+10", -1.75e+10)]
    [TestCase("+1.75E+10", 1.75e+10)]
    [TestCase("2E+308", null)]
    [TestCase("-2E+308", null)]
    [TestCase("number", null)]
    public void GetOptionalDouble_reads_xsd_compliant_double_values(string xmlText, double? expectedValue)
    {
        // Generally speaking, uint is stored as int in the internal representation, because nearly
        // all API expects int and it is just so much easier to work with.
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalDouble(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue));
    }

    [TestCase("2025-10-25", "2025-10-25T00:00:00")]
    [TestCase("2004-04-12T13:20:00Z", "2004-04-12T13:20:00Z")]
    [TestCase("today", null)]
    public void GetOptionalDateTime_reads_xsd_compliant_dateTime_values(string xmlText, string expectedString)
    {
        DateTimeOffset? expectedValue = expectedString is not null ? DateTimeOffset.Parse(expectedString) : null;
        using var reader = CreateReader(xmlText);
        var readValue = reader.GetOptionalDateTime(AttributeName);

        Assert.That(readValue, Is.EqualTo(expectedValue?.DateTime));
    }

    [Test]
    public void GetOptionalString_returns_stored_string_without_XString_decoding()
    {
        // 0x9 (tab) is invalid character per XML 1.0.
        // 0x57 (W) is a valid character per XML 1.0.
        const string value = @"Dear &lt;user_name&gt; _x0009_ - _x0057_elcome";

        using var reader = CreateReader(value);
        var readValue = reader.GetOptionalString(AttributeName);

        Assert.That(readValue, Is.EqualTo(@"Dear <user_name> _x0009_ - _x0057_elcome"));
    }

    [TestCase("def", BindingFlags.Default)]
    [TestCase("ci", BindingFlags.IgnoreCase)]
    [TestCase("CI", null)] // Enums names are case sensitive
    [TestCase("Default", null)] // name is not matched, only configured string
    [TestCase("", null)]
    [TestCase("NonExpectedValue", null)]
    public void GetOptionalEnum_returns_enum_parsed_by_enum_mapper(string xmlText, BindingFlags? enumValue)
    {
        var mapper = new XmlToEnumMapper.Builder().Add(new BiDictionary<string, BindingFlags>
        {
            { "def", BindingFlags.Default },
            { "ci", BindingFlags.IgnoreCase },
        }).Build();

        using var reader = CreateReader(xmlText, mapper);
        var readValue = reader.GetOptionalEnum<BindingFlags>(AttributeName);

        Assert.That(readValue, Is.EqualTo(enumValue));
    }

    private static XmlTreeReader CreateReader(string attributeValue, XmlToEnumMapper mapper = null)
    {
        var xmlContext = $"<element {AttributeName}=\"{attributeValue}\"/>";
        var xmlReader = XmlReader.Create(new StringReader(xmlContext));
        mapper ??= new XmlToEnumMapper.Builder().Build();
        var reader = new XmlTreeReader(xmlReader, mapper, false);
        reader.Open("element", string.Empty);
        return reader;
    }
}
