using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests;

[TestFixture]
public class XLHelperTests
{
    [Test]
    public void IsValidColumnTest()
    {
        Assert.Multiple(() =>
        {
            Assert.That(XLHelper.IsValidColumn(""), Is.False);
            Assert.That(XLHelper.IsValidColumn("1"), Is.False);
            Assert.That(XLHelper.IsValidColumn("A1"), Is.False);
            Assert.That(XLHelper.IsValidColumn("AA1"), Is.False);
            Assert.That(XLHelper.IsValidColumn("A"), Is.True);
            Assert.That(XLHelper.IsValidColumn("AA"), Is.True);
            Assert.That(XLHelper.IsValidColumn("AAA"), Is.True);
            Assert.That(XLHelper.IsValidColumn("Z"), Is.True);
            Assert.That(XLHelper.IsValidColumn("ZZ"), Is.True);
            Assert.That(XLHelper.IsValidColumn("XFD"), Is.True);
            Assert.That(XLHelper.IsValidColumn("ZAA"), Is.False);
            Assert.That(XLHelper.IsValidColumn("XZA"), Is.False);
            Assert.That(XLHelper.IsValidColumn("XFZ"), Is.False);
        });
    }

    [Test]
    public void ReplaceRelative1()
    {
        var result = XLHelper.ReplaceRelative("A1", 2, "B");
        Assert.That(result, Is.EqualTo("B2"));
    }

    [Test]
    public void ReplaceRelative2()
    {
        var result = XLHelper.ReplaceRelative("$A1", 2, "B");
        Assert.That(result, Is.EqualTo("$A2"));
    }

    [Test]
    public void ReplaceRelative3()
    {
        var result = XLHelper.ReplaceRelative("A$1", 2, "B");
        Assert.That(result, Is.EqualTo("B$1"));
    }

    [Test]
    public void ReplaceRelative4()
    {
        var result = XLHelper.ReplaceRelative("$A$1", 2, "B");
        Assert.That(result, Is.EqualTo("$A$1"));
    }

    [Test]
    public void ReplaceRelative5()
    {
        var result = XLHelper.ReplaceRelative("1:1", 2, "B");
        Assert.That(result, Is.EqualTo("2:2"));
    }

    [Test]
    public void ReplaceRelative6()
    {
        var result = XLHelper.ReplaceRelative("$1:1", 2, "B");
        Assert.That(result, Is.EqualTo("$1:2"));
    }

    [Test]
    public void ReplaceRelative7()
    {
        var result = XLHelper.ReplaceRelative("1:$1", 2, "B");
        Assert.That(result, Is.EqualTo("2:$1"));
    }

    [Test]
    public void ReplaceRelative8()
    {
        var result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
        Assert.That(result, Is.EqualTo("$1:$1"));
    }

    [Test]
    public void ReplaceRelative9()
    {
        var result = XLHelper.ReplaceRelative("A:A", 2, "B");
        Assert.That(result, Is.EqualTo("B:B"));
    }

    [Test]
    public void ReplaceRelativeA()
    {
        var result = XLHelper.ReplaceRelative("$A:A", 2, "B");
        Assert.That(result, Is.EqualTo("$A:B"));
    }

    [Test]
    public void ReplaceRelativeB()
    {
        var result = XLHelper.ReplaceRelative("A:$A", 2, "B");
        Assert.That(result, Is.EqualTo("B:$A"));
    }

    [Test]
    public void ReplaceRelativeC()
    {
        var result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
        Assert.That(result, Is.EqualTo("$A:$A"));
    }

    [TestCase("Sheet1", "Sheet1")]
    [TestCase("O'Brien's sales", "O'Brien's sales")]
    [TestCase(" data # ", " data # ")]
    [TestCase("data $1.00", "data $1.00")]
    [TestCase("data ", "data?")]
    [TestCase("abc def", "abc/def")]
    [TestCase("data 0 ", "data[0]")]
    [TestCase("data ", "data*")]
    [TestCase("abc def", "abc\\def")]
    [TestCase(" data", "'data")]
    [TestCase("data ", "data'")]
    [TestCase("d'at'a", "d'at'a")]
    [TestCase("sheet a4", "sheet:a4")]
    [TestCase("null", null)]
    [TestCase("empty", "")]
    [TestCase("1234567890123456789012345678901", "1234567890123456789012345678901TOOLONG")]
    public void CreateSafeSheetNames(string expected, string input)
    {
        var actual = XLHelper.CreateSafeSheetName(input);
        Assert.That(actual, Is.EqualTo(expected));
    }

    [TestCase("Sheet1", ExpectedResult = "Sheet1")]
    [TestCase("O'Brien's sales", ExpectedResult = "O'Brien's sales")]
    [TestCase(" data # ", ExpectedResult = " data # ")]
    [TestCase("data $1.00", ExpectedResult = "data $1.00")]
    [TestCase("data?", ExpectedResult = "data_")]
    [TestCase("abc/def", ExpectedResult = "abc_def")]
    [TestCase("data[0]", ExpectedResult = "data_0_")]
    [TestCase("data*", ExpectedResult = "data_")]
    [TestCase("abc\\def", ExpectedResult = "abc_def")]
    [TestCase("'data", ExpectedResult = "_data")]
    [TestCase("data'", ExpectedResult = "data_")]
    [TestCase("d'at'a", ExpectedResult = "d'at'a")]
    [TestCase("sheet:a4", ExpectedResult = "sheet_a4")]
    [TestCase(null, ExpectedResult = "null")]
    [TestCase("", ExpectedResult = "empty")]
    [TestCase("1234567890123456789012345678901TOOLONG", ExpectedResult = "1234567890123456789012345678901")]
    public string CreateSafeSheetNamesWithUnderscore(string input)
    {
        return XLHelper.CreateSafeSheetName(input, replaceChar: '_');
    }

    [Test]
    public void CreateSafeSheetNamesInvalidReplacementChar()
    {
        Assert.Throws<ArgumentException>(() => XLHelper.CreateSafeSheetName("abc\\def", replaceChar: ':'));
    }
}