namespace ClosedXML.Excel;

internal readonly record struct XLFillKey
{
    public required XLColorKey BackgroundColor { get; init; }

    public required XLColorKey PatternColor { get; init; }

    public required XLFillPatternValues PatternType { get; init; }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = 0;

            if (!HasNoFill())
            {
                hash = (int)PatternType;
                hash = (hash * 397) ^ BackgroundColor.GetHashCode();

                if (!HasNoForeground())
                {
                    hash = (hash * 397) ^ PatternColor.GetHashCode();
                }
            }

            return hash;
        }
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