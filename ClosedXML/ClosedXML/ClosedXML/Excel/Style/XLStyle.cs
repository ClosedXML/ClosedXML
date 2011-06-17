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
                Protection = new XLProtection(container, initialStyle.Protection);
            }
            else
            {
                Font = new XLFont(container, null);
                Alignment = new XLAlignment(container);
                Border = new XLBorder(container, null);
                Fill = new XLFill(container);
                NumberFormat = new XLNumberFormat(container, null);
                Protection = new XLProtection(container, null);
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
            sb.Append(Font.ToString());
            sb.Append(" Fill:");
            sb.Append(Fill.ToString());
            sb.Append(" Border:");
            sb.Append(Border.ToString());
            sb.Append(" NumberFormat: ");
            sb.Append(NumberFormat.ToString());
            sb.Append(" Alignment: ");
            sb.Append(Alignment.ToString());
            sb.Append(" Protection: ");
            sb.Append(Protection.ToString());
            return sb.ToString();
        }

        public bool Equals(IXLStyle other)
        {
            return
            //    this.Font.GetHashCode().Equals(other.Font.GetHashCode()) 
            //&&  this.Fill.GetHashCode().Equals(other.Fill.GetHashCode())
            //&& this.Border.GetHashCode().Equals(other.Border.GetHashCode())
            //&& this.NumberFormat.GetHashCode().Equals(other.NumberFormat.GetHashCode())
            //&& this.Alignment.GetHashCode().Equals(other.Alignment.GetHashCode())
            //&& this.Protection.GetHashCode().Equals(other.Protection.GetHashCode()) 

                this.Font.Equals(other.Font)
            &&  this.Fill.Equals(other.Fill)
            &&  this.Border.Equals(other.Border)
            &&  this.NumberFormat.Equals(other.NumberFormat)
            &&  this.Alignment.Equals(other.Alignment)
            &&  this.Protection.Equals(other.Protection)
            ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLStyle)obj);
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
