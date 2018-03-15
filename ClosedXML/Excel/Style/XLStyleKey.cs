using System;

namespace ClosedXML.Excel
{
    public struct XLStyleKey : IEquatable<XLStyleKey>
    {
        public XLAlignmentKey Alignment { get; set; }

        public XLBorderKey Border { get; set; }

        public XLFillKey Fill { get; set; }

        public XLFontKey Font { get; set; }

        public XLNumberFormatKey NumberFormat { get; set; }

        public XLProtectionKey Protection { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -476701294;
            hashCode = hashCode * -1521134295 + Alignment.GetHashCode();
            hashCode = hashCode * -1521134295 + Border.GetHashCode();
            hashCode = hashCode * -1521134295 + Fill.GetHashCode();
            hashCode = hashCode * -1521134295 + Font.GetHashCode();
            hashCode = hashCode * -1521134295 + NumberFormat.GetHashCode();
            hashCode = hashCode * -1521134295 + Protection.GetHashCode();
            return hashCode;
        }

        public bool Equals(XLStyleKey other)
        {
            return Alignment == other.Alignment &&
                   Border == other.Border &&
                   Fill == other.Fill &&
                   Font == other.Font &&
                   NumberFormat == other.NumberFormat &&
                   Protection == other.Protection;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLStyleKey)
                return Equals((XLStyleKey)obj);
            return base.Equals(obj);
        }

        public static bool operator ==(XLStyleKey left, XLStyleKey right) => left.Equals(right);

        public static bool operator !=(XLStyleKey left, XLStyleKey right) => !(left.Equals(right));
    }
}
