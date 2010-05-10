using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    
    public class XLCell
    {
        public XLCellAddress CellAddress { get; private set; }
        private XLWorkbook workbook;
        public XLCell(XLWorkbook workbook, XLCellAddress cellAddress)
        {
            this.CellAddress = cellAddress;
            this.workbook = workbook;
        }

        public XLCell(XLWorkbook workbook, String cellAddressString)
        {
            this.CellAddress = new XLCellAddress(cellAddressString);
            this.workbook = workbook;
        }

        public UInt32 Row { get { return CellAddress.Row; } }
        public UInt32 Column { get { return CellAddress.Column; } }
        public String ColumnLetter { get { return XLWorksheet.ColumnNumberToLetter(this.Column); } }

        public CellValues DataType { get; private set; }
        public String InnerValue { get; private set; }
        public String Value {
            get
            {
                if (DataType == CellValues.Boolean)
                {
                    return (InnerValue == "1").ToString();
                }
                else if (DataType == CellValues.SharedString)
                {
                   return workbook.SharedStrings.GetString(UInt32.Parse(InnerValue));
                }
                else if (DataType == CellValues.Date)
                {
                    return GetDateTimeFromSerial(Int32.Parse(InnerValue)).ToString();
                }
                else
                {
                    return InnerValue;
                }
            }
            set
            {
                String val = value;

                Double dTest;
                DateTime dtTest;
                Boolean bTest;
                if (Double.TryParse(val, out dTest))
                {
                    DataType = CellValues.Number;
                }
                else if (DateTime.TryParse(val, out dtTest))
                {
                    DataType = CellValues.Date;
                    String datePart = GetSerialFromDateTime(dtTest.Day, dtTest.Month, dtTest.Year).ToString();
                    val = datePart;
                }
                else if (Boolean.TryParse(val, out bTest))
                {
                    DataType = CellValues.Boolean;
                    val = bTest ? "1" : "0";
                }
                else
                {
                    DataType = CellValues.SharedString;
                    val = workbook.SharedStrings.Add(val).ToString();
                }
                InnerValue = val;
                HasValue = !value.Equals(String.Empty);
            }
        }

        public Boolean HasValue { get; private set; }

        private DateTime GetDateTimeFromSerial(Int32 SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }

        private Int32 GetSerialFromDateTime(Int32 nDay, Int32 nMonth, Int32 nYear)
        {
            // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a

            // leap year, but Excel/Lotus 123 think it is...

            if (nDay == 29 && nMonth == 02 && nYear == 1900)
                return 60;

            // DMY to Modified Julian calculatie with an extra substraction of 2415019.

            long nSerialDate =
                    (int)((1461 * (nYear + 4800 + (int)((nMonth - 14) / 12))) / 4) +
                    (int)((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
                    (int)((3 * ((int)((nYear + 4900 + (int)((nMonth - 14) / 12)) / 100))) / 4) +
                    nDay - 2415019 - 32075;

            if (nSerialDate < 60)
            {
                // Because of the 29-02-1900 bug, any serial date 

                // under 60 is one off... Compensate.

                nSerialDate--;
            }

            return (int)nSerialDate;
        }


    }
}
