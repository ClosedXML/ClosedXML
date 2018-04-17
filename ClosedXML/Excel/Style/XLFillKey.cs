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
            hashCode = hashCode * -1521134295 + PatternType.GetHashCode();
            return hashCode;
        }

        public bool Equals(XLFillKey other)
        {
            return
                (PatternType == XLFillPatternValues.None && other.PatternType == XLFillPatternValues.None) ||
                BackgroundColor == other.BackgroundColor
             && PatternColor == other.PatternColor
             && PatternType == other.PatternType;
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
