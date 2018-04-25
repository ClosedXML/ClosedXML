using System;

namespace ClosedXML.Excel
{
    [System.Diagnostics.DebuggerDisplay("{RangeType} {RangeAddress}")]
    internal struct XLRangeKey : IEquatable<XLRangeKey>
    {
        public XLRangeType RangeType { get; private set; }

        public XLRangeAddress RangeAddress { get; private set; }

        public XLRangeKey(XLRangeType rangeType, XLRangeAddress address)
        {
            RangeType = rangeType;
            RangeAddress = address;
        }

        #region Overrides

        /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
        /// <returns>true if the current object is equal to the <paramref name="other" /> parameter; otherwise, false.</returns>
        /// <param name="other">An object to compare with this object.</param>
        public bool Equals(XLRangeKey other)
        {
            return RangeType == other.RangeType &&
                   RangeAddress.Equals(other.RangeAddress);
        }

        /// <summary>Indicates whether this instance and a specified object are equal.</summary>
        /// <returns>true if <paramref name="obj" /> and this instance are the same type and represent the same value; otherwise, false.</returns>
        /// <param name="obj">Another object to compare to. </param>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is XLRangeKey && Equals((XLRangeKey)obj);
        }

        /// <summary>Returns the hash code for this instance.</summary>
        /// <returns>A 32-bit signed integer that is the hash code for this instance.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)RangeType * 397) ^ RangeAddress.GetHashCode();
            }
        }

        public static bool operator ==(XLRangeKey left, XLRangeKey right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(XLRangeKey left, XLRangeKey right)
        {
            return !(left == right);
        }

        #endregion Overrides
    }
}
