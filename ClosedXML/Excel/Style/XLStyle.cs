using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStyle : IXLStyle
    {
        public XLStyle(IXLStylized container, IXLStyle initialStyle = null, Boolean useDefaultModify = true)
        {
            if (initialStyle != null)
            {
                Font = new XLFont(container, initialStyle.Font, useDefaultModify);
                Alignment = new XLAlignment(container, initialStyle.Alignment);
                Border = new XLBorder(container, initialStyle.Border, useDefaultModify);
                Fill = new XLFill(container, initialStyle.Fill, useDefaultModify);
                NumberFormat = new XLNumberFormat(container, initialStyle.NumberFormat);
                Protection = new XLProtection(container, initialStyle.Protection);
            }
            else
            {
                Font = new XLFont(container, null);
                Alignment = new XLAlignment(container);
                Border = new XLBorder(container, null);
                Fill = new XLFill(container);
                NumberFormat = new XLNumberFormat(container, null);
                Protection = new XLProtection(container);
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

        public IXLProtection Protection { get; set; }

        public IXLNumberFormat DateFormat { get; private set; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Font:");
            sb.Append(Font);
            sb.Append(" Fill:");
            sb.Append(Fill);
            sb.Append(" Border:");
            sb.Append(Border);
            sb.Append(" NumberFormat: ");
            sb.Append(NumberFormat);
            sb.Append(" Alignment: ");
            sb.Append(Alignment);
            sb.Append(" Protection: ");
            sb.Append(Protection);
            return sb.ToString();
        }

        public bool Equals(IXLStyle other)
        {
            return
                Font.Equals(other.Font)
            &&  Fill.Equals(other.Fill)
            &&  Border.Equals(other.Border)
            &&  NumberFormat.Equals(other.NumberFormat)
            &&  Alignment.Equals(other.Alignment)
            &&  Protection.Equals(other.Protection)
            ;
        }

        public override bool Equals(object obj)
        {
            return Equals((XLStyle)obj);
        }

        public override int GetHashCode()
        {
            return Font.GetHashCode()
                ^ Fill.GetHashCode()
                ^ Border.GetHashCode()
                ^ NumberFormat.GetHashCode()
                ^ Alignment.GetHashCode()
                ^ Protection.GetHashCode();
        }
    }
}
