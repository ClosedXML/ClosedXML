using NUnit.Framework;
using ClosedXML.IO;

namespace ClosedXML.Tests.IO;

internal class XStringConvertTests
{
    [TestCase("", "")]
    [TestCase("_x000D_", "\r")]
    [TestCase("_x30ab_", "カ")] // Hexadecimal numbers are case insensitive
    [TestCase("_x0009_", "\t")]
    [TestCase("__x0041__", "_A_")]
    [TestCase("A_x0042_C", "ABC")]
    [TestCase("_X0041_", "_X0041_")] // Must be lowercase x in the pattern
    [TestCase("_x263A_", "\u263a")] // Smiley face
    [TestCase("_xD83D__xDE43_", "\ud83d\ude43")] // Astral planes - Upside down smiley face
    [TestCase("Result:_x0009_ _x0057_", "Result:\t W")]
    [TestCase("DE_x005F_xAB50_0161_title", "DE_xAB50_0161_title")]
    [TestCase("_x0001_ _x0002_ _x0003_ _x0004_", "\u0001 \u0002 \u0003 \u0004")]
    [TestCase("_x0005_ _x0006_ _x0007_ _x0008_", "\u0005 \u0006 \u0007 \u0008")]
    [TestCase("_xaaBB_ _xAAbb_", "\uAABB \uAABB")]
    [TestCase(@"_Xceed_Something", @"_Xceed_Something")] // https://github.com/ClosedXML/ClosedXML/issues/1154
    [TestCase("_xD83DDE43_", "_xD83DDE43_")] // 8 hex digit name, decoded by XmlConvert.DecodeName, but not by XString
    public void Decodes_encoded_unicode_characters(string sourceText, string expectedText)
    {
        var decodedText = XStringConvert.Decode(sourceText);

        Assert.That(decodedText, Is.EqualTo(expectedText));
    }
}
