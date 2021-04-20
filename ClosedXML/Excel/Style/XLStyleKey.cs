using System;

namespace ClosedXML.Excel
{
    using DocumentFormat.OpenXml.Office2010.PowerPoint;

    internal struct XLStyleKey : IEquatable<XLStyleKey>
    {
        private XLAlignmentKey _alignment;

        private XLBorderKey _border;

        private XLFillKey _fill;

        private XLFontKey _font;

        private Boolean _includeQuotePrefix;

        private XLNumberFormatKey _numberFormat;

        private XLProtectionKey _protection;

        private int _cachedHashCode;

        public XLAlignmentKey Alignment
        {
            get { return _alignment; }
            set
            {
                _alignment = value;
                _cachedHashCode = 0;
            }
        }

        public XLBorderKey Border
        {
            get { return _border; }
            set
            {
                _border = value;
                _cachedHashCode = 0;
            }
        }

        public XLFillKey Fill
        {
            get { return _fill; }
            set
            {
                _fill = value;
                _cachedHashCode = 0;
            }
        }

        public XLFontKey Font
        {
            get { return _font; }
            set
            {
                _font = value;
                _cachedHashCode = 0;
            }
        }

        public Boolean IncludeQuotePrefix
        {
            get { return _includeQuotePrefix; }
            set
            {
                _includeQuotePrefix = value;
                _cachedHashCode = 0;
            }
        }

        public XLNumberFormatKey NumberFormat
        {
            get { return _numberFormat; }
            set
            {
                _numberFormat = value;
                _cachedHashCode = 0;
            }
        }

        public XLProtectionKey Protection
        {
            get { return _protection; }
            set
            {
                _protection = value;
                _cachedHashCode = 0;
            }
        }

        public override int GetHashCode()
        {
            if (_cachedHashCode != 0)
            {
                return _cachedHashCode;
            }

            var hashCode = -476701294;
            hashCode = hashCode * -1521134295 + Alignment.GetHashCode();
            hashCode = hashCode * -1521134295 + Border.GetHashCode();
            hashCode = hashCode * -1521134295 + Fill.GetHashCode();
            hashCode = hashCode * -1521134295 + Font.GetHashCode();
            hashCode = hashCode * -1521134295 + IncludeQuotePrefix.GetHashCode();
            hashCode = hashCode * -1521134295 + NumberFormat.GetHashCode();
            hashCode = hashCode * -1521134295 + Protection.GetHashCode();

            if (hashCode == 0) hashCode = 1;
            _cachedHashCode = hashCode;

            return hashCode;
        }

        public bool Equals(XLStyleKey other)
        {
            return Alignment == other.Alignment &&
                   Border == other.Border &&
                   Fill == other.Fill &&
                   Font == other.Font &&
                   IncludeQuotePrefix == other.IncludeQuotePrefix &&
                   NumberFormat == other.NumberFormat &&
                   Protection == other.Protection;
        }

        public override string ToString()
        {
            return
                this == XLStyle.Default.Key ? "Default" : 
                string.Format("Alignment: {0} Border: {1} Fill: {2} Font: {3} IncludeQuotePrefix: {4} NumberFormat: {5} Protection: {6}",
                    Alignment == XLStyle.Default.Key.Alignment ? "Default" : Alignment.ToString(),
                    Border == XLStyle.Default.Key.Border ? "Default" : Border.ToString(),
                    Fill == XLStyle.Default.Key.Fill ? "Default" : Fill.ToString(),
                    Font == XLStyle.Default.Key.Font ? "Default" : Font.ToString(),
                    IncludeQuotePrefix == XLStyle.Default.Key.IncludeQuotePrefix ? "Default" : IncludeQuotePrefix.ToString(),
                    NumberFormat == XLStyle.Default.Key.NumberFormat ? "Default" : NumberFormat.ToString(),
                    Protection == XLStyle.Default.Key.Protection ? "Default" : Protection.ToString());
        }

        public override bool Equals(object obj)
        {
            if (obj is XLStyleKey)
                return Equals((XLStyleKey)obj);
            return base.Equals(obj);
        }

        public static bool operator ==(XLStyleKey left, XLStyleKey right) => left.Equals(right);

        public static bool operator !=(XLStyleKey left, XLStyleKey right) => !(left.Equals(right));

        public void Deconstruct(
            out XLAlignmentKey alignment,
            out XLBorderKey border,
            out XLFillKey fill,
            out XLFontKey font,
            out Boolean includeQuotePrefix,
            out XLNumberFormatKey numberFormat,
            out XLProtectionKey protection)
        {
            alignment = Alignment;
            border = Border;
            fill = Fill;
            font = Font;
            includeQuotePrefix = IncludeQuotePrefix;
            numberFormat = NumberFormat;
            protection = Protection;
        }
    }
}
