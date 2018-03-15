using System;

namespace ClosedXML.Excel
{
    public struct XLBorderKey : IEquatable<XLBorderKey>
    {
        public XLBorderStyleValues LeftBorder { get; set; }

        public XLColorKey LeftBorderColor { get; set; }

        public XLBorderStyleValues RightBorder { get; set; }

        public XLColorKey RightBorderColor { get; set; }

        public XLBorderStyleValues TopBorder { get; set; }

        public XLColorKey TopBorderColor { get; set; }

        public XLBorderStyleValues BottomBorder { get; set; }

        public XLColorKey BottomBorderColor { get; set; }

        public XLBorderStyleValues DiagonalBorder { get; set; }

        public XLColorKey DiagonalBorderColor { get; set; }

        public bool DiagonalUp { get; set; }

        public bool DiagonalDown { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -198124310;
            hashCode = hashCode * -1521134295 + LeftBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + LeftBorderColor.GetHashCode();
            hashCode = hashCode * -1521134295 + RightBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + RightBorderColor.GetHashCode();
            hashCode = hashCode * -1521134295 + TopBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + TopBorderColor.GetHashCode();
            hashCode = hashCode * -1521134295 + BottomBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + BottomBorderColor.GetHashCode();
            hashCode = hashCode * -1521134295 + DiagonalBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + DiagonalBorderColor.GetHashCode();
            hashCode = hashCode * -1521134295 + DiagonalUp.GetHashCode();
            hashCode = hashCode * -1521134295 + DiagonalDown.GetHashCode();
            return hashCode;
        }

        public bool Equals(XLBorderKey other)
        {
            return
               LeftBorder == other.LeftBorder
            && LeftBorderColor == other.LeftBorderColor
            && RightBorder == other.RightBorder
            && RightBorderColor == other.RightBorderColor
            && TopBorder == other.TopBorder
            && TopBorderColor == other.TopBorderColor
            && BottomBorder == other.BottomBorder
            && BottomBorderColor == other.BottomBorderColor
            && DiagonalBorder == other.DiagonalBorder
            && DiagonalBorderColor == other.DiagonalBorderColor
            && DiagonalUp == other.DiagonalUp
            && DiagonalDown == other.DiagonalDown;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLBorderKey)
                return Equals((XLBorderKey)obj);
            return base.Equals(obj);
        }

        public static bool operator ==(XLBorderKey left, XLBorderKey right) => left.Equals(right);

        public static bool operator !=(XLBorderKey left, XLBorderKey right) => !(left.Equals(right));
    }
}
