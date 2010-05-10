using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLWorksheet: XLRange
    {
        public const UInt32 MaxNumberOfRows = 1048576;
        public const UInt32 MaxNumberOfColumns = 16384;

        public XLWorksheet(XLWorkbook workbook, String sheetName, XLCells cells)
            : base(
                  new XLCell(workbook, new XLCellAddress(1,1))
                , new XLCell(workbook, new XLCellAddress(MaxNumberOfRows, MaxNumberOfColumns))
                , cells, null)
        {
            this.name = sheetName;
        }

        public override List<XLCell> Cells()
        {
            return Cells(CellContent.All);
        }

        public override List<XLCell> Cells(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                return base.Cells(cellContent);
            }
            else
            {
                String errorText = "Cannot load entire worksheet into memory. Please use a range (eg. XLWorksheet.Range(\"A1:D5\"))  or retrieve cells with values (eg. XLWorksheet.Cells(CellContent.WithValues)).";
                throw new InvalidOperationException(errorText);
            }
        }
        
        private String name;
        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        public static String ColumnNumberToLetter(UInt32 column)
        {
            String s = String.Empty;
            for (
                Int32 i = Convert.ToInt32(
                    Math.Log(
                        Convert.ToDouble(
                            25 * (
                                Convert.ToDouble(column)
                                + 1
                            )
                         )
                     ) / Math.Log(26)
                 ) - 1
                ; i >= 0
                ; i--
                )
            {
                Int32 x = Convert.ToInt32(Math.Pow(26, i + 1) - 1) / 25 - 1;
                if (column > x)
                {
                    s += (Char)(((column - x - 1) / Convert.ToInt32(Math.Pow(26, i))) % 26 + 65);
                }
            }
            return s;
        }

        public static UInt32 ColumnLetterToNumber(String column)
        {
            Int32 intColumnLetterLength = column.Length;
            Int32 retVal = 0;
            for (Int32 intCount = 0; intCount < intColumnLetterLength; intCount++)
            {
                retVal = retVal * 26 + (column.Substring(intCount, 1).ToUpper().ToCharArray()[0] - 64);
            }
            return (UInt32)retVal;
        }
    }
}
