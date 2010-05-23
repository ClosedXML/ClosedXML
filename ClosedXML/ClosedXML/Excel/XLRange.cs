using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public enum XLCellValues { SharedString, Number, Boolean, DateTime }

    public class XLRange
    {
        private Dictionary<XLAddress, XLRange> worksheetCells;
        public Boolean IsWorksheet { get; private set; }
        public Boolean IsCell { get; private set; }
        private XLRange Worksheet { get; set; }
        public XLRange ParentRange { get; private set; }
        private String worksheetName;
        public String Name { get { return ToString(); } }
        private Boolean isInitializing;
        private XLWorkbook workbook;
        public XLRange(XLAddress firstCellAddress, XLAddress lastCellAddress, XLRange parentRange, XLRange worksheet, String name = "", XLWorkbook workbook=null)
        {
            isInitializing = true;
            this.worksheetName = name;
            this.workbook = workbook;
            FirstCellAddress = firstCellAddress;
            LastCellAddress = lastCellAddress;
            IsCell = firstCellAddress == lastCellAddress;

            ParentRange = parentRange;
            if (worksheet != null)
            {
                IsWorksheet = false;
                Worksheet = worksheet;
                CellStyle = new XLStyle(parentRange.CellStyle, this);
            }
            else
            {
                IsWorksheet = true;
                worksheetCells = new Dictionary<XLAddress, XLRange>();
                Worksheet = this;
                var defaultStyle = new XLStyle(workbook.WorkbookStyle, this);

                CellStyle = defaultStyle;
            }
            isInitializing = false;
        }

        public XLAddress Address { get { return FirstCellAddress; } }

        private XLAddress FirstCellAddress { get; set; }
        public XLRange FirstCell
        {
            get
            {
                return Cell(FirstCellAddress);
            }
        }

        private XLAddress LastCellAddress { get; set; }
        public XLRange LastCell
        {
            get
            {
                return Cell(LastCellAddress);
            }
        }

        public XLAddress AddressInWorksheet { get { return FirstCellAddressInWorksheet; } }

        private XLAddress FirstCellAddressInWorksheet
        {
            get
            {
                if (IsWorksheet || ParentRange.IsWorksheet)
                    return FirstCellAddress;
                else
                    return FirstCellAddress + ParentRange.FirstCellAddressInWorksheet - 1;
            }
        }

        private XLAddress LastCellAddressAbsolute
        {
            get
            {
                if (IsWorksheet || ParentRange.IsWorksheet)
                    return LastCellAddress;
                else
                    return ParentRange.FirstCellAddressInWorksheet + LastCellAddress - 1;
            }
        }

        public XLRange Range(XLAddress firstCellAddress, XLAddress lastCellAddress)
        {
            if (lastCellAddress > (LastCellAddress - FirstCellAddress + 1))
                throw new IndexOutOfRangeException("Cell addresses must be within parent range.");

            return new XLRange(firstCellAddress, lastCellAddress, this, Worksheet);
        }

        public XLRange Range(String firstCellAddress, String lastCellAddress)
        {
            return Range(new XLAddress(firstCellAddress), new XLAddress(lastCellAddress));
        }

        public XLRange Cell(XLAddress address)
        {
            XLAddress absoluteCellAddress;
            if (IsWorksheet || ParentRange.IsWorksheet)
                absoluteCellAddress = address;
            else
                absoluteCellAddress = ParentRange.FirstCellAddressInWorksheet + address - 1;

            var cell = Worksheet.GetCell(absoluteCellAddress);
            return cell;
        }

        public XLRange Cell(String addressString)
        {
            var cellAddress = new XLAddress(addressString);
            return Cell(cellAddress);
        }

        public XLRange Cell(UInt32 row, UInt32 column)
        {
            var cellAddress = new XLAddress(row, column);
            return Cell(cellAddress);
        }

        protected XLRange GetCell(XLAddress cellAddress)
        {
            if (!worksheetCells.ContainsKey(cellAddress))
            {
                worksheetCells.Add(cellAddress, new XLRange(cellAddress, cellAddress, this, this));
            }

            return worksheetCells[cellAddress];
        }

        public Boolean HasValue { get; private set; }

        public XLCellValues DataType { get; set; }

        private String cellValue;
        public String Value 
        {
            get
            {
                return Worksheet.GetCell(FirstCellAddressInWorksheet).cellValue;
            }
            set
            {
                String val = value;

                Double dTest;
                DateTime dtTest;
                Boolean bTest;
                if (Double.TryParse(val, out dTest))
                {
                    DataType = XLCellValues.Number;
                }
                else if (DateTime.TryParse(val, out dtTest))
                {
                    DataType = XLCellValues.DateTime;
                    String dateSerial = GetSerialFromDateTime(dtTest).ToString();
                    Style.NumberFormat.NumberFormatId = 14;
                    val = dateSerial;
                }
                else if (Boolean.TryParse(val, out bTest))
                {
                    DataType = XLCellValues.Boolean;
                    val = bTest ? "1" : "0";
                }
                else
                {
                    DataType = XLCellValues.SharedString;
                }

                Worksheet.GetCell(FirstCellAddressInWorksheet).cellValue = val;
                Worksheet.GetCell(FirstCellAddressInWorksheet).HasValue = !val.Equals(String.Empty);
            }
        }


        internal XLStyle CellStyle;
        public XLStyle Style
        {
            get
            {
                CellStyle = new XLStyle(Worksheet.GetCell(FirstCellAddressInWorksheet).CellStyle, this);
                return CellStyle;
            }
            set
            {
                Cells().ForEach(c => c.CellStyle = new XLStyle(value, this));
            }
        }

        public IEnumerable<XLRange> Cells()
        {
            if (IsWorksheet)
            {
                foreach (var cell in worksheetCells.Where(c=>c.Key != LastCellAddress).Select(c=>c.Value))
                {
                    yield return cell;
                }
            }
            else
            {
                for (UInt32 row = FirstCellAddressInWorksheet.Row; row <= LastCellAddressAbsolute.Row; row++)
                {
                    for (UInt32 column = FirstCellAddressInWorksheet.Column; column <= LastCellAddressAbsolute.Column; column++)
                    {
                        yield return Worksheet.GetCell(new XLAddress(row, column));
                    }
                }
            }
        }

        private DateTime GetDateTimeFromSerial(Double serialDate)
        {
            String sDate = serialDate.ToString();
            Int32 datePart = Int32.Parse(sDate.Split('.').First());
            Double timePart = serialDate - (Double)datePart;
            if (datePart > 59) datePart -= 1; //Excel/Lotus 2/29/1900 bug   

            TimeSpan maxTime = new TimeSpan(0, 23, 59, 59, 999);
            Double totalMilliseconds = maxTime.TotalMilliseconds * timePart;

            return new DateTime(1899, 12, 31).AddDays(datePart).AddMilliseconds(totalMilliseconds);
        }

        private Double GetSerialFromDateTime(DateTime dateTime)
        {
            Int32 nDay = dateTime.Day;
            Int32 nMonth = dateTime.Month;
            Int32 nYear = dateTime.Year;

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

            TimeSpan timePart = new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second);
            TimeSpan maxTime = new TimeSpan(0, 23, 59, 59, 999);
            Double timeRatio = timePart.TotalMilliseconds / maxTime.TotalMilliseconds;

            return (Double)nSerialDate + timeRatio;
        }



        internal void ProcessCells(Action<XLRange> action)
        {
            if (!(IsWorksheet || isInitializing)) //(IsCell && ParentRange.IsWorksheet)))
            {
                Cells().ForEach(c => action(c));
            }  
        }

        #region Overridden

        public override string ToString()
        {
            if (IsWorksheet)
                return worksheetName;
            else
                return FirstCellAddress.ToString();
        }

        #endregion

    }
}
