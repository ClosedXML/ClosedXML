using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLCell: IXLCell
    {
        public XLCell(IXLAddress address, IXLStyle defaultStyle)
        {
            this.Address = address;
            Style = defaultStyle;
            if (Style == null) Style = XLWorkbook.DefaultStyle;
        }
        public Boolean Initialized { get; private set; }
        public IXLAddress Address { get; private set; }

        private String cellValue;
        public String Value
        {
            get
            {
                return cellValue;
            }
            set
            {
                String val = value;

                Double dTest;
                //DateTime dtTest;
                Boolean bTest;
                if (Double.TryParse(val, out dTest))
                {
                    DataType = XLCellValues.Number;
                }
                //else if (DateTime.TryParse(val, out dtTest))
                //{
                //    DataType = XLCellValues.DateTime;
                //    Style.NumberFormat = new OPNumberFormat(14);
                //    val = dtTest.ToOADate().ToString();
                //}
                else if (Boolean.TryParse(val, out bTest))
                {
                    DataType = XLCellValues.Boolean;
                    val = bTest ? "1" : "0";
                }
                else
                {
                    DataType = XLCellValues.SharedString;
                }
                cellValue = val;
            }
        }

        #region IXLStylized Members

        private IXLStyle style;
        public IXLStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = new XLStyle(null, value);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get 
            {
                UpdatingStyle = true;
                yield return style;
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion


        public XLCellValues DataType { get; set; }
    }
}
