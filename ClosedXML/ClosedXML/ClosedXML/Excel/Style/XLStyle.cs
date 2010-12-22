using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStyle : IXLStyle
    {
        public XLStyle(IXLStylized container, IXLStyle initialStyle = null)
        {
            if (initialStyle != null)
            {
                Font = new XLFont(container, initialStyle.Font);
                Alignment = new XLAlignment(container, initialStyle.Alignment);
                Border = new XLBorder(container, initialStyle.Border);
                Fill = new XLFill(container, initialStyle.Fill);
                NumberFormat = new XLNumberFormat(container, initialStyle.NumberFormat);
            }
            else
            {
                Font = new XLFont(container, null);
                Alignment = new XLAlignment(container);
                Border = new XLBorder(container, null);
                Fill = new XLFill(container);
                NumberFormat = new XLNumberFormat(container, null);
            }

            DateFormat = NumberFormat;
        }

        public IXLFont Font { get; set; }

        public IXLAlignment Alignment { get; set; }

        public IXLBorder Border { get; set; }

        public IXLFill Fill { get; set; }

        private IXLNumberFormat numberFormat;
        public IXLNumberFormat NumberFormat 
        {
            get
            {
                return numberFormat;
            }
            set
            {
                numberFormat = value;
                DateFormat = numberFormat;
            }
        }

        public IXLNumberFormat DateFormat { get; private set; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Font:");
            sb.Append(Font.ToString());
            sb.Append(" Fill:");
            sb.Append(Fill.ToString());
            sb.Append(" Border:");
            sb.Append(Border.ToString());
            sb.Append(" NumberFormat: ");
            sb.Append(NumberFormat.ToString());
            sb.Append(" Alignment: ");
            sb.Append(Alignment.ToString());
            return sb.ToString();
        }

        public bool Equals(IXLStyle other)
        {
            return
            this.Font.Equals(other.Font)
            && this.Fill.Equals(other.Fill)
            && this.Border.Equals(other.Border)
            && this.NumberFormat.Equals(other.NumberFormat)
            && this.Alignment.Equals(other.Alignment)
            ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLStyle)obj);
        }

        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                int hash = 17;
                hash = hash * 23 + Font.GetHashCode();
                hash = hash * 23 + Fill.GetHashCode();
                hash = hash * 23 + Border.GetHashCode();
                hash = hash * 23 + NumberFormat.GetHashCode();
                hash = hash * 23 + Alignment.GetHashCode();
                return hash;
            }
        }
    }
}
