using System;

namespace ClosedXML.Excel
{
    internal struct XLStyleKey : IEquatable<XLStyleKey>
    {
        public XLAlignmentKey Alignment { get; set; }

        public XLBorderKey Border { get; set; }

        public XLFillKey Fill { get; set; }

        public XLFontKey Font { get; set; }

        public Boolean IncludeQuotePrefix { get; set; }

        public string Name { get; set; }

        public XLNumberFormatKey NumberFormat { get; set; }

        public XLProtectionKey Protection { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -476701294;
            hashCode = hashCode * -1521134295 + Alignment.GetHashCode();
            hashCode = hashCode * -1521134295 + Border.GetHashCode();
            hashCode = hashCode * -1521134295 + Fill.GetHashCode();
            hashCode = hashCode * -1521134295 + Font.GetHashCode();
            hashCode = hashCode * -1521134295 + IncludeQuotePrefix.GetHashCode();
            if (Name != null)
                hashCode = hashCode * -1521134295 + StringComparer.InvariantCultureIgnoreCase.GetHashCode(Name);
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
                   IncludeQuotePrefix == other.IncludeQuotePrefix &&
                   StringComparer.InvariantCultureIgnoreCase.Equals(Name, other.Name) &&
                   NumberFormat == other.NumberFormat &&
                   Protection == other.Protection;
        }

        public override string ToString()
        {
            return
                this == XLStyle.Default.Key ? "Default" : 
                string.Format("Alignment: {0} Border: {1} Fill: {2} Font: {3} IncludeQuotePrefix: {4} Name:{5} NumberFormat: {6} Protection: {7}",
                    Alignment == XLStyle.Default.Key.Alignment ? "Default" : Alignment.ToString(),
                    Border == XLStyle.Default.Key.Border ? "Default" : Border.ToString(),
                    Fill == XLStyle.Default.Key.Fill ? "Default" : Fill.ToString(),
                    Font == XLStyle.Default.Key.Font ? "Default" : Font.ToString(),
                    IncludeQuotePrefix == XLStyle.Default.Key.IncludeQuotePrefix ? "Default" : IncludeQuotePrefix.ToString(),
                    Name ?? XLStyle.Default.Value.Name,
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
