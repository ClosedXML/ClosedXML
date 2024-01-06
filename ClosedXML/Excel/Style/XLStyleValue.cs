using ClosedXML.Excel.Caching;
using System;

namespace ClosedXML.Excel
{
    internal sealed class XLStyleValue : IEquatable<XLStyleValue?>
    {
        private static readonly XLStyleRepository Repository = new(key => new XLStyleValue(key));
        private int? _hashCode; // Cached hash key

        public static XLStyleValue FromKey(ref XLStyleKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLStyleKey DefaultKey = new()
        {
            Alignment = XLAlignmentValue.Default.Key,
            Border = XLBorderValue.Default.Key,
            Fill = XLFillValue.Default.Key,
            Font = XLFontValue.Default.Key,
            IncludeQuotePrefix = false,
            NumberFormat = XLNumberFormatValue.Default.Key,
            Protection = XLProtectionValue.Default.Key
        };

        internal static readonly XLStyleValue Default = FromKey(ref DefaultKey);

        public XLStyleKey Key { get; }

        public XLAlignmentValue Alignment { get; }

        public XLBorderValue Border { get; }

        public XLFillValue Fill { get; }

        public XLFontValue Font { get; }

        public Boolean IncludeQuotePrefix { get; }

        public XLNumberFormatValue NumberFormat { get; }

        public XLProtectionValue Protection { get; }

        internal XLStyleValue(XLStyleKey key)
        {
            Key = key;
            var (alignment, border, fill, font, _, numberFormat, protection) = Key;
            Alignment = XLAlignmentValue.FromKey(ref alignment);
            Border = XLBorderValue.FromKey(ref border);
            Fill = XLFillValue.FromKey(ref fill);
            Font = XLFontValue.FromKey(ref font);
            IncludeQuotePrefix = key.IncludeQuotePrefix;
            NumberFormat = XLNumberFormatValue.FromKey(ref numberFormat);
            Protection = XLProtectionValue.FromKey(ref protection);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            return Equals(obj as XLStyleValue);
        }

        public bool Equals(XLStyleValue? other)
        {
            if (other is null)
                return false;

            if (_hashCode.HasValue && other._hashCode.HasValue && _hashCode != other._hashCode)
                return false;

            return Key.Equals(other.Key);
        }

        public override int GetHashCode()
        {
            if (_hashCode.HasValue)
                return _hashCode.Value;

            _hashCode = -280332839 + Key.GetHashCode();
            return _hashCode.Value;
        }

        public static bool operator ==(XLStyleValue? left, XLStyleValue? right)
        {
            if (left is null)
                return right is null;

            return left.Equals(right);
        }

        public static bool operator !=(XLStyleValue? left, XLStyleValue? right)
        {
            return !(left == right);
        }

        internal XLStyleValue WithAlignment(Func<XLAlignmentValue, XLAlignmentValue> modify)
        {
            return WithAlignment(modify(Alignment));
        }

        internal XLStyleValue WithAlignment(XLAlignmentValue alignment)
        {
            var keyCopy = Key;
            keyCopy.Alignment = alignment.Key;
            return FromKey(ref keyCopy);
        }

        internal XLStyleValue WithIncludeQuotePrefix(bool includeQuotePrefix)
        {
            var keyCopy = Key;
            keyCopy.IncludeQuotePrefix = includeQuotePrefix;
            return FromKey(ref keyCopy);
        }

        internal XLStyleValue WithNumberFormat(XLNumberFormatValue numberFormat)
        {
            var keyCopy = Key;
            keyCopy.NumberFormat = numberFormat.Key;
            return FromKey(ref keyCopy);
        }
    }
}
