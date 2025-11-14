using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles;

public class ProtectionTests
{
    [Test]
    [TestCaseSource(nameof(ProtectionApiSetters))]
    public void Protection_property_can_be_individually_set(FormatTestCase<IXLProtection> testCase)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var cellFormat = ws.Cell("B2").Style;
        foreach (var testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellFormat.Protection, testValue);
            var setValue = testCase.GetPropertyValue(cellFormat.Protection);
            Assert.AreEqual(testValue, setValue);
        }
    }

    [Test]
    public void Protection_can_be_copied()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var targetProtection = ws.Cell("A2").Style.Protection;
        Assert.IsTrue(targetProtection.Locked);
        Assert.IsFalse(targetProtection.Hidden);
        var source = ws.Cell("A1").Style
            .Protection.SetLocked(false)
            .Protection.SetHidden(true);

        targetProtection = source.Protection;

        Assert.IsFalse(targetProtection.Locked);
        Assert.IsTrue(targetProtection.Hidden);
    }

    [Test]
    public void Protection_has_equality_comparison()
    {
        Action<IXLProtection>[] changePropertyToNonDefault =
        {
            x => x.SetLocked(false),
            x => x.SetHidden(true)
        };

        using var wb = new XLWorkbook();
        foreach (var changeProperty in changePropertyToNonDefault)
        {
            var ws = wb.AddWorksheet();
            var lhs = ws.Cell("A1").Style.Protection;
            var rhs = ws.Cell("A2").Style.Protection;

            Assert.AreEqual(lhs, rhs);
            changeProperty(lhs);
            Assert.AreNotEqual(lhs, rhs);
        }
    }

    private static IEnumerable<object> ProtectionApiSetters()
    {
        var boolValues = new[] { false, true };
        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Hidden, (protection, value) => protection.Hidden = value, boolValues);
        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Hidden, (protection, value) => protection.SetHidden(value), boolValues);
        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Hidden, (protection, _) => protection.SetHidden(), true);

        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Locked, (protection, value) => protection.Locked = value, boolValues);
        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Locked, (protection, value) => protection.SetLocked(value), boolValues);
        yield return FormatTestCase<IXLProtection>.ForProtection(protection => protection.Locked, (protection, _) => protection.SetLocked(), true);
    }
}
