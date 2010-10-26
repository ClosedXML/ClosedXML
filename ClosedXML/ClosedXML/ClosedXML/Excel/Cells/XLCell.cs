using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLCell : IXLCell
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
                Double dTest;
                DateTime dtTest;
                Boolean bTest;
                if (initialized)
                {
                    if (dataType == XLCellValues.Boolean)
                    {
                        if (Boolean.TryParse(val, out bTest))
                            val = bTest ? "1" : "0";
                        else if (!(val == "1" || val == "0"))
                            throw new ArgumentException("'" + val + "' is not a Boolean type.");
                    }
                    else if (dataType == XLCellValues.DateTime)
                    {
                        if (DateTime.TryParse(val, out dtTest))
                        {

                            val = dtTest.ToOADate().ToString();
                        }
                        else if (!Double.TryParse(val, out dTest))
                        {
                            throw new ArgumentException("'" + val + "' is not a DateTime type.");
                        }

                        if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                            Style.NumberFormat.NumberFormatId = 14;
                    }
                    else if (dataType == XLCellValues.Number)
                    {
                        if (!Double.TryParse(val, out dTest))
                            throw new ArgumentException("'" + val + "' is not a Numeric type.");
                        
                    }
                }
                else
                {
                    if (val.Length > 0 && val.Substring(0, 1) == "'")
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
                if (cellValue.Length > 0)
                {
                    if (value == XLCellValues.Boolean)
                    {
                        Boolean bTest;
                        if (Boolean.TryParse(cellValue, out bTest))
                            cellValue = Boolean.Parse(cellValue) ? "1" : "0";
                        else
                            cellValue = value != 0 ? "1" : "0";
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

                        if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                            Style.NumberFormat.NumberFormatId = 14;
                    }
                    else if (value == XLCellValues.Number)
                    {
                        cellValue = Double.Parse(cellValue).ToString();
                        //if (Style.NumberFormat.Format == String.Empty )
                        //    Style.NumberFormat.NumberFormatId = 0;
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
                }
                dataType = value;
            }
        }
    }
}
