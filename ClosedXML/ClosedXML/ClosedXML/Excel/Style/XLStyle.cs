using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OX.Copyable;

namespace ClosedXML.Excel.Style
{
    public class XLStyle: IXLStyle
    {
        public XLStyle(IXLStylized container, IXLStyle initialStyle = null)
        {
            if (initialStyle != null)
            {
                //Font = new XLFont(container, initialStyle.Font);
                Font = (IXLFont)initialStyle.Font.Copy();
                Alignment = (IXLAlignment)initialStyle.Alignment.Copy();
                Border = (IXLBorder)initialStyle.Border.Copy();
                Fill = (IXLFill)initialStyle.Fill.Copy();
                NumberFormat = (IXLNumberFormat)initialStyle.NumberFormat.Copy();
            }
            else
            {
                //Font = new XLFont(container);
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
