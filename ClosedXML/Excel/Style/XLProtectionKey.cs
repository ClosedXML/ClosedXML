using System;

namespace ClosedXML.Excel
{
    internal struct XLProtectionKey : IEquatable<XLProtectionKey>
    {
        public bool Locked { get; set; }

        public bool Hidden { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -1357408252;
            hashCode = hashCode * -1521134295 + Locked.GetHashCode();
            hashCode = hashCode * -1521134295 + Hidden.GetHashCode();
            return hashCode;
        }

        public bool Equals(XLProtectionKey other)
        {
            return
                Locked == other.Locked
             && Hidden == other.Hidden;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLProtectionKey)
                return Equals((XLProtectionKey)obj);
            return base.Equals(obj);
        }

        public override string ToString()
        {
            return (Locked ? "Locked" : "") + (Hidden ? "Hidden" : "");
        }

        public static bool operator ==(XLProtectionKey left, XLProtectionKey right) => left.Equals(right);

        public static bool operator !=(XLProtectionKey left, XLProtectionKey right) => !(left.Equals(right));
    }
}
