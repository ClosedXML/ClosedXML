using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal class XLStyleValue
    {
        private static readonly XLStyleRepository Repository = new XLStyleRepository(key => new XLStyleValue(key));

        public static XLStyleValue FromKey(ref XLStyleKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLStyleKey DefaultKey = new XLStyleKey
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

        public XLStyleKey Key { get; private set; }

        public XLAlignmentValue Alignment { get; private set; }

        public XLBorderValue Border { get; private set; }

        public XLFillValue Fill { get; private set; }

        public XLFontValue Font { get; private set; }

        public bool IncludeQuotePrefix { get; private set; }

        public XLNumberFormatValue NumberFormat { get; private set; }

        public XLProtectionValue Protection { get; private set; }

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
            {
                return true;
            }

            var cached = obj as XLStyleValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            if (_hashCode.HasValue)
            {
                return _hashCode.Value;
            }

            _hashCode = -280332839 + Key.GetHashCode();
            return _hashCode.Value;
        }

        public static bool operator ==(XLStyleValue left, XLStyleValue right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (left is null && right is null)
            {
                return true;
            }

            if (left is null || right is null)
            {
                return false;
            }

            if (left._hashCode.HasValue && right._hashCode.HasValue &&
                left._hashCode != right._hashCode)
            {
                return false;
            }

            return left.Key.Equals(right.Key);
        }

        public static bool operator !=(XLStyleValue left, XLStyleValue right)
        {
            return !(left == right);
        }

        private int? _hashCode;
    }
}
