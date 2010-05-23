using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public class XLNumberFormat
    {
        #region Properties

        private XLRange range;

        private UInt32? numberFormatId;
        public UInt32? NumberFormatId
        {
            get { return numberFormatId; }
            set
            {
                numberFormatId = value;
                format = String.Empty;
                if (range != null) range.ProcessCells(c =>
                {
                    c.CellStyle.NumberFormat.numberFormatId = value;
                    c.CellStyle.NumberFormat.format = String.Empty;
                });
            }
        }

        private String format = String.Empty;
        public String Format
        {
            get { return format; }
            set
            {
                format = value;
                numberFormatId = null;
                if (range != null) range.ProcessCells(c =>
                {
                    c.CellStyle.NumberFormat.numberFormatId = null;
                    c.CellStyle.NumberFormat.format = value;
                });
            }
        }

        #endregion

        #region Constructors

        public XLNumberFormat(XLNumberFormat numberFormat, XLRange range)
        {
            this.range = range;
            if (numberFormat != null)
            {
                if (numberFormat.numberFormatId.HasValue)
                    NumberFormatId = numberFormat.NumberFormatId.Value;
                else
                    Format = numberFormat.Format;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            String numberFormatIdString = numberFormatId.HasValue ? numberFormatId.Value.ToString() : "n/a";
            return numberFormatIdString + "-" + format.ToString();
        }

        #endregion
    }
}
