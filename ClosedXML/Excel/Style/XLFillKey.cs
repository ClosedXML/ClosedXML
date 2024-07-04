using System;

namespace ClosedXML.Excel;

internal readonly record struct XLFillKey
{
    public required XLColorKey BackgroundColor { get; init; }

    public required XLColorKey PatternColor { get; init; }

    public required XLFillPatternValues PatternType { get; init; }

    public override int GetHashCode()
    {
        var hash = new HashCode();

        if (HasNoFill()) return hash.ToHashCode();

        hash.Add(PatternType);
        hash.Add(BackgroundColor);

        if (HasNoForeground()) return hash.ToHashCode();
                
        hash.Add(PatternColor);
            
        return hash.ToHashCode();
    }

    public bool Equals(XLFillKey other)
    {
        if (HasNoFill() && other.HasNoFill())
            return true;

        return BackgroundColor == other.BackgroundColor
               && PatternType == other.PatternType
               && (HasNoForeground() && other.HasNoForeground() ||
                   PatternColor == other.PatternColor);
    }

    private bool HasNoFill()
    {
        return PatternType == XLFillPatternValues.None
               || (PatternType == XLFillPatternValues.Solid && XLColor.IsTransparent(BackgroundColor));
    }

    private bool HasNoForeground()
    {
        return PatternType == XLFillPatternValues.Solid ||
               PatternType == XLFillPatternValues.None;
    }

    public override string ToString()
    {
        return $"{PatternType} {BackgroundColor}/{PatternColor}";
    }
}
