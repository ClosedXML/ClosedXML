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
    public void Decodes_encoded_unicode_characters(string sourceText, string expectedText)
    {
        var decodedText = XStringConvert.Decode(sourceText);

        Assert.That(decodedText, Is.EqualTo(expectedText));
    }
}
