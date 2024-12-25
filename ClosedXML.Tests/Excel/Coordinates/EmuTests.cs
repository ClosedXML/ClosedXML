using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Coordinates;

[TestFixture]
internal class EmuTests
{
    [TestCase(0.14, AbsLengthUnit.Inch, 128_016)]
    [TestCase(2.43, AbsLengthUnit.Centimeter, 874_800)]
    [TestCase(748, AbsLengthUnit.Millimeter, 26_928_000)]
    [TestCase(23.9, AbsLengthUnit.Point, 303_530)]
    [TestCase(4.157, AbsLengthUnit.Pica, 633_527)]
    [TestCase(14.6, AbsLengthUnit.Emu, 15)]
    [TestCase(2348.52, AbsLengthUnit.Inch, null)]
    public void From_converts_value_to_emu(double value, AbsLengthUnit unit, int? emu)
    {
        Assert.AreEqual(emu, Emu.From(value, unit)?.Value);
    }

    [TestCase(AbsLengthUnit.Inch, 5.9912904636920388)]
    [TestCase(AbsLengthUnit.Centimeter, 15.217877777777778)]
    [TestCase(AbsLengthUnit.Millimeter, 152.17877777777778)]
    [TestCase(AbsLengthUnit.Point, 431.3729133858268)]
    [TestCase(AbsLengthUnit.Pica, 35.94774278215223)]
    [TestCase(AbsLengthUnit.Emu, 5_478_436)]
    public void To_converts_to_specified_unit(AbsLengthUnit unit, double value)
    {
        Assert.AreEqual(value, Emu.From(5_478_436, AbsLengthUnit.Emu)?.To(unit));
    }

    [Test]
    [SetCulture("cs-CZ")]
    public void ToString_uses_culture_invariant_format()
    {
        Assert.AreEqual("1.4mm", Emu.From(1.4, AbsLengthUnit.Millimeter).ToString());
    }
}
