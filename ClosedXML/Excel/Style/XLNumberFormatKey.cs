#nullable disable

using System;

namespace ClosedXML.Excel
{
    internal struct XLNumberFormatKey : IEquatable<XLNumberFormatKey>
    {
        /// <summary>
        /// Number format identifier of predefined format, see <see cref="XLPredefinedFormat"/>.
        /// If -1, the format is custom and stored in the <see cref="Format"/>.
        /// </summary>
        public int NumberFormatId { get; set; }

        public string Format { get; set; }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = -759193072;
                hashCode = hashCode * -1521134295 + NumberFormatId;
                hashCode = hashCode * -1521134295 + (Format == null ? 0 : Format.GetHashCode());
                return hashCode;
            }
        }

        public bool Equals(XLNumberFormatKey other)
        {
            return NumberFormatId == other.NumberFormatId &&
                   string.Equals(Format, other.Format);
        }

        public override bool Equals(object obj)
        {
            return obj is XLNumberFormatKey other && Equals(other);
        }

        public override string ToString()
        {
            return $"{Format}/{NumberFormatId}";
        }

        public static bool operator ==(XLNumberFormatKey left, XLNumberFormatKey right) => left.Equals(right);

        public static bool operator !=(XLNumberFormatKey left, XLNumberFormatKey right) => !(left.Equals(right));
    }
}
