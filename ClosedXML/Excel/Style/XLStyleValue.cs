using ClosedXML.Excel.Caching;
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An immutable style value.
    /// </summary>
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

        internal XLStyleKey Key { get; }

        internal XLAlignmentValue Alignment { get; }

        internal XLBorderValue Border { get; }

        internal XLFillValue Fill { get; }

        internal XLFontValue Font { get; }

        internal Boolean IncludeQuotePrefix { get; }

        internal XLNumberFormatValue NumberFormat { get; }

        internal XLProtectionValue Protection { get; }

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

        /// <summary>
        /// Combine row and column styles into a combined style. This style is used by non-pinged
        /// cells of a worksheet.
        /// </summary>
        internal static XLStyleValue Combine(XLStyleValue sheetStyle, XLStyleValue rowStyle, XLStyleValue colStyle)
        {
            var isRowSame = ReferenceEquals(sheetStyle, rowStyle);
            var isColSame = ReferenceEquals(sheetStyle, colStyle);

            if (isRowSame && isColSame)
                return sheetStyle;

            // At least one style is different, maybe both.
            if (isRowSame)
                return colStyle;

            if (isColSame)
                return rowStyle;

            // Both styles are different from sheet one, merge. If both style components differ,
            // row has a preference because Excel gives it preference. Generally, if there is
            // row/col style conflict, all cells affected by conflict should be materialized (aka
            // 'pinged') during row/col style modification and have their own style explicitly
            // specified to avoid ambiguity, so we shouldn't really need to rely on this
            // resolution, it's just last ditch effort.
            var alignment = GetExplicitlySet(sheetStyle.Alignment, rowStyle.Alignment, colStyle.Alignment);
            var border = GetExplicitlySet(sheetStyle.Border, rowStyle.Border, colStyle.Border);
            var fill = GetExplicitlySet(sheetStyle.Fill, rowStyle.Fill, colStyle.Fill);
            var font = GetExplicitlySet(sheetStyle.Font, rowStyle.Font, colStyle.Font);
            var includeQuotePrefix = GetExplicitlySet(sheetStyle.IncludeQuotePrefix, rowStyle.IncludeQuotePrefix, colStyle.IncludeQuotePrefix);
            var numberFormat = GetExplicitlySet(sheetStyle.NumberFormat, rowStyle.NumberFormat, colStyle.NumberFormat);
            var protection = GetExplicitlySet(sheetStyle.Protection, rowStyle.Protection, colStyle.Protection);

            var combinedStyleKey = new XLStyleKey
            {
                Alignment = alignment.Key,
                Border = border.Key,
                Fill = fill.Key,
                Font = font.Key,
                IncludeQuotePrefix = includeQuotePrefix,
                NumberFormat = numberFormat.Key,
                Protection = protection.Key,
            };
            return Repository.GetOrCreate(ref combinedStyleKey);

            static T GetExplicitlySet<T>(T sheetComponent, T rowComponent, T colComponent)
                where T : notnull
            {
                // Use reference equal to speed up the process instead of standard equals.
                var rowHasSameComponent = typeof(T).IsClass
                    ? ReferenceEquals(sheetComponent, rowComponent)
                    : sheetComponent.Equals(rowComponent);
                var colHasSameComponent = typeof(T).IsClass
                    ? ReferenceEquals(sheetComponent, colComponent)
                    : sheetComponent.Equals(colComponent);

                if (rowHasSameComponent && colHasSameComponent)
                    return sheetComponent;

                // At least one style is different, maybe both.
                if (rowHasSameComponent)
                    return colComponent;

                // If col has same component as sheet, we should return row.
                // If both are different, row component should have precedence.
                return rowComponent;
            }
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
