using System;

namespace ClosedXML.Excel
{
    internal struct XLFillKey : IEquatable<XLFillKey>
    {
        public XLColorKey BackgroundColor { get; set; }

        public XLColorKey PatternColor { get; set; }

        public XLFillPatternValues PatternType { get; set; }

        public override int GetHashCode()
        {
            var hashCode = 2043579837;
            hashCode = hashCode * -1521134295 + BackgroundColor.GetHashCode();
            hashCode = hashCode * -1521134295 + PatternColor.GetHashCode();

            var patternType = PatternType;
            if (BackgroundColor.ColorType == XLColorType.Indexed && BackgroundColor.Indexed == 64)
                patternType = XLFillPatternValues.None;

            hashCode = hashCode * -1521134295 + patternType.GetHashCode();

            return hashCode;
        }

        public bool Equals(XLFillKey other)
        {
            if (PatternType == XLFillPatternValues.None && other.PatternType == XLFillPatternValues.None)
                return true;

            var patternType1 = PatternType;
            var patternType2 = other.PatternType;

            if (BackgroundColor.ColorType == XLColorType.Indexed && BackgroundColor.Indexed == 64)
                patternType1 = XLFillPatternValues.None;

            if (other.BackgroundColor.ColorType == XLColorType.Indexed && other.BackgroundColor.Indexed == 64)
                patternType2 = XLFillPatternValues.None;

            return BackgroundColor == other.BackgroundColor
                   && PatternColor == other.PatternColor
                   && patternType1 == patternType2;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLFillKey)
                return Equals((XLFillKey)obj);
            return base.Equals(obj);
        }

        public override string ToString()
        {
            return $"{PatternType} {BackgroundColor}/{PatternColor}";
        }

        public static bool operator ==(XLFillKey left, XLFillKey right) => left.Equals(right);

        public static bool operator !=(XLFillKey left, XLFillKey right) => !(left.Equals(right));
    }
}
