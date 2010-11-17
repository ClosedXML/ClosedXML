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
                Font = new XLFont(container);
                Alignment = new XLAlignment(container);
                Border = new XLBorder(container);
                Fill = new XLFill(container);
                NumberFormat = new XLNumberFormat(container);
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

    }
}
