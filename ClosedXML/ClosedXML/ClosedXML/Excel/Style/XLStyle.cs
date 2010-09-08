using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLStyle: IXLStyle
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
        }

        public IXLFont Font { get; set; }

        public IXLAlignment Alignment { get; set; }

        public IXLBorder Border { get; set; }

        public IXLFill Fill { get; set; }

        public IXLNumberFormat NumberFormat { get; set; }

        public override string ToString()
        {
            return
                "Font:" + Font.ToString()
                + " Fill:" + Fill.ToString()
                + " Border:" + Border.ToString()
                + " NumberFormat: " + NumberFormat.ToString()
                + " Alignment: " + Alignment.ToString();
        }
    }
}
