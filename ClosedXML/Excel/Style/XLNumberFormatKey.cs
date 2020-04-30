using System;

namespace ClosedXML.Excel
{
    internal struct XLNumberFormatKey : IEquatable<XLNumberFormatKey>
    {
        public int NumberFormatId { get; set; }

        public string Format { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -759193072;
            hashCode = hashCode * -1521134295 + NumberFormatId;
            hashCode = hashCode * -1521134295 + (Format == null ? 0 : Format.GetHashCode());
            return hashCode;
        }

        public bool Equals(XLNumberFormatKey other)
        {
            return
                NumberFormatId == other.NumberFormatId
             && string.Equals(Format, other.Format);
        }

        public override bool Equals(object obj)
        {
            if (obj is XLNumberFormatKey)
                return Equals((XLNumberFormatKey)obj);
            return base.Equals(obj);
        }

        public override string ToString()
        {
            return $"{Format}/{NumberFormatId}";
        }

        public static bool operator ==(XLNumberFormatKey left, XLNumberFormatKey right) => left.Equals(right);

        public static bool operator !=(XLNumberFormatKey left, XLNumberFormatKey right) => !(left.Equals(right));
    }
}
