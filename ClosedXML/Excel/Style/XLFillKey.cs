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

            if (HasNoFill()) return hashCode;

            hashCode = hashCode * -1521134295 + (int)PatternType;
            hashCode = hashCode * -1521134295 + BackgroundColor.GetHashCode();

            if (HasNoForeground()) return hashCode;
                
            hashCode = hashCode * -1521134295 + PatternColor.GetHashCode();
            
            return hashCode;
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
