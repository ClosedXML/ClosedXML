using System;

namespace ClosedXML.Excel;

internal sealed partial class XLFillCellFormat : IXLFill
{
    XLColor IXLFill.BackgroundColor
    {
        get => BackgroundColor;
        set => BackgroundColor = value;
    }

    XLColor IXLFill.PatternColor
    {
        get => PatternColor;
        set => PatternColor = value;
    }

    XLFillPatternValues IXLFill.PatternType
    {
        get => PatternType;
        set => PatternType = value;
    }

    IXLStyle IXLFill.SetBackgroundColor(XLColor value)
    {
        BackgroundColor = value;
        return _parent;
    }

    IXLStyle IXLFill.SetPatternColor(XLColor value)
    {
        PatternColor = value;
        return _parent;
    }

    IXLStyle IXLFill.SetPatternType(XLFillPatternValues value)
    {
        PatternType = value;
        return _parent;
    }

    bool IEquatable<IXLFill>.Equals(IXLFill other)
    {
        // This is a "business" equality, i.e. will both fills look the same.
        // This is gradient fill, other can only represent pattern, regardless of what it actually is.
        var isPatternFill = _parent.Resolve(static x => x.Fill.Pattern) is not null;
        if (!isPatternFill)
            return false;

        if (!HasFill(this) && !HasFill(other))
            return true;

        if (PatternType != other.PatternType)
            return false;

        if (BackgroundColor != other.BackgroundColor)
            return false;

        var patternColor = UsesPatternColor(this) ? PatternColor : null;
        var otherPatternColor = UsesPatternColor(other) ? other.PatternColor : null;
        if (patternColor != otherPatternColor)
            return false;

        return true;

        static bool HasFill(IXLFill fill)
        {
            var patternType = fill.PatternType;
            if (patternType == XLFillPatternValues.None)
                return false;

            if (patternType == XLFillPatternValues.Solid && fill.BackgroundColor.IsAuto)
                return false;

            return true;
        }

        static bool UsesPatternColor(IXLFill fill)
        {
            return fill.PatternType is not XLFillPatternValues.None and not XLFillPatternValues.Solid;
        }
    }
}
