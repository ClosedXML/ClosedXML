using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLCell : IXLCell
    {
        public XLCell(IXLAddress address, IXLStyle defaultStyle)
        {
            this.Address = address;
            Style = defaultStyle;
            if (Style == null) Style = XLWorkbook.DefaultStyle;
        }

        public IXLAddress Address { get; private set; }

        private Boolean initialized = false;
        private String cellValue = String.Empty;
        public String Value
        {
            get
            {
                return cellValue;
            }
            set
            {
                String val = value;

                if (!initialized)
                {
                    Double dTest;
                    DateTime dtTest;
                    Boolean bTest;
                    if (val.Substring(0, 1) == "'")
                    {
                        val = val.Substring(1, val.Length - 1);
                        dataType = XLCellValues.Text;
                    }
                    else if (Double.TryParse(val, out dTest))
                    {
                        dataType = XLCellValues.Number;
                    }
                    else if (DateTime.TryParse(val, out dtTest))
                    {
                        dataType = XLCellValues.DateTime;
                        Style.NumberFormat.NumberFormatId = 14;
                        val = dtTest.ToOADate().ToString();
                    }
                    else if (Boolean.TryParse(val, out bTest))
                    {
                        dataType = XLCellValues.Boolean;
                        val = bTest ? "1" : "0";
                    }
                    else
                    {
                        dataType = XLCellValues.Text;
                    }
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

        private XLCellValues dataType;
        public XLCellValues DataType
        {
            get
            {
                return dataType;
            }
            set
            {
                initialized = true;
                if (value == XLCellValues.Boolean)
                {
                    cellValue = Boolean.Parse(cellValue) ? "1" : "0";
                }
                else if (value == XLCellValues.DateTime)
                {
                    DateTime dtTest;
                    Double dblTest;
                    if (DateTime.TryParse(cellValue, out dtTest))
                    {
                        cellValue = dtTest.ToOADate().ToString();
                    }
                    else if (Double.TryParse(cellValue, out dblTest))
                    {
                        cellValue = dblTest.ToString();
                    }
                    else
                    {
                        throw new ArgumentException("Cannot set data type to DateTime because '" + cellValue + "' is not recognized as a date.");
                    }

                    if (Style.NumberFormat.Format == String.Empty)
                        Style.NumberFormat.NumberFormatId = 14;
                }
                else if (value == XLCellValues.Number)
                {
                    cellValue = Double.Parse(cellValue).ToString();
                    if (Style.NumberFormat.Format == String.Empty)
                        Style.NumberFormat.NumberFormatId = 0;
                }
                else
                {
                    if (dataType == XLCellValues.Boolean)
                    {
                        cellValue = (cellValue == "0" ? false : true).ToString();
                    }
                    else if (dataType == XLCellValues.Number)
                    {
                        cellValue = Double.Parse(cellValue).ToString(Style.NumberFormat.Format);
                    }
                    else if (dataType == XLCellValues.DateTime)
                    {
                        cellValue = DateTime.FromOADate(Double.Parse(cellValue)).ToString(Style.NumberFormat.Format);
                    }
                }

                dataType = value;
            }
        }
    }
}
