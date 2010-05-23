using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.Style
{
    public class XLStyle
    {
        XLRange range;
        public XLStyle(XLStyle initialStyle, XLRange range)
        {
            this.range = range;
            if (initialStyle != null)
            {
                Font = new XLFont(initialStyle.Font, range);
                Fill = new XLFill(initialStyle.Fill, range);
                Border = new XLBorder(initialStyle.Border, range);
                NumberFormat = new XLNumberFormat(initialStyle.NumberFormat, range);
                Alignment = new XLAlignment(initialStyle.Alignment, range);
            }
            else
            {
                Font = new XLFont(null, range);
                Fill = new XLFill(null, range);
                Border = new XLBorder(null, range);
                NumberFormat = new XLNumberFormat(null, range);
                Alignment = new XLAlignment(null, range);
            }
        }

        public XLFont Font { get; set; }
        public XLFill Fill { get; set; }
        public XLBorder Border { get; set; }
        public XLNumberFormat NumberFormat { get; set; }
        public XLAlignment Alignment { get; set; }

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