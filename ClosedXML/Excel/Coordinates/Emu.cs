using System;
using System.Diagnostics;
using System.Globalization;

namespace ClosedXML.Excel;

/// <summary>
/// English metric unit.
/// </summary>
internal readonly record struct Emu
{
    private const int PerInch = 914400;
    private const int PerCm = 360000;
    private const int PerMm = PerCm / 10;
    private const int PerPt = PerInch / 72;
    private const int PerPc = PerInch / 6;

    private readonly AbsLengthUnit _preferredUnit;

    internal static readonly Emu ZeroPt = new(0, AbsLengthUnit.Point);

    private Emu(int emus, AbsLengthUnit preferredUnit)
    {
        _preferredUnit = preferredUnit;
        Value = emus;
    }

    /// <summary>
    /// Length in EMU.
    /// </summary>
    internal int Value { get; }

    internal static Emu? From(double value, AbsLengthUnit srcUnit)
    {
        var coef = GetUnitCoefficient(srcUnit);
        var emus = Math.Round(value * coef, MidpointRounding.AwayFromZero);
        if (emus is < int.MinValue or > int.MaxValue)
            return null;

        return new Emu((int)emus, srcUnit);
    }

    private static int GetUnitCoefficient(AbsLengthUnit srcUnit)
    {
        return srcUnit switch
        {
            AbsLengthUnit.Inch => PerInch,
            AbsLengthUnit.Centimeter => PerCm,
            AbsLengthUnit.Millimeter => PerMm,
            AbsLengthUnit.Point => PerPt,
            AbsLengthUnit.Pica => PerPc,
            AbsLengthUnit.Emu => 1,
            _ => throw new ArgumentOutOfRangeException(),
        };
    }

    /// <summary>
    /// Return length in specified unit.
    /// </summary>
    internal double To(AbsLengthUnit unit)
    {
        var coef = GetUnitCoefficient(unit);
        return Value / (double)coef;
    }

    public override string ToString()
    {
        return ToString(_preferredUnit);
    }

    public string ToString(AbsLengthUnit unit)
    {
        var lengthInUnit = To(unit);
        var unitSuffix = unit switch
        {
            AbsLengthUnit.Inch => "in",
            AbsLengthUnit.Centimeter => "cm",
            AbsLengthUnit.Millimeter => "mm",
            AbsLengthUnit.Point => "pt",
            AbsLengthUnit.Pica => "pc",
            AbsLengthUnit.Emu => "emu",
            _ => throw new UnreachableException(),
        };
        return lengthInUnit.ToString(CultureInfo.InvariantCulture) + unitSuffix;
    }
}
