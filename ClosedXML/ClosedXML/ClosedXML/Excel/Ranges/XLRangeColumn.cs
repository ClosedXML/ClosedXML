using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRangeColumn: XLRangeBase, IXLRangeColumn
    {
        public XLRangeColumn(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters.RangeAddress)
        {
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(this.RangeAddress, range, columnsShifted);
        }
        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(this.RangeAddress, range, rowsShifted);
        }

        public IXLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        public IXLCells Cells(String cellsInColumn)
        {
            var retVal = new XLCells(Worksheet, false, false, false);
            var rangePairs = cellsInColumn.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }
            return retVal;
        }

        public IXLCells Cells(Int32 firstRow, Int32 lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }
        
        public IXLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
        }
        public override IXLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                String[] arrRange = rangeAddressStr.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(rangeAddressToUse);
            return Range(rangeAddress);
        }

        public void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }
        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return InsertCellsAbove(numberOfRows, false);
        }
        public IXLCells InsertCellsAbove(int numberOfRows, Boolean expandRange)
        {
            return InsertRowsAbove(numberOfRows, expandRange).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return InsertCellsBelow(numberOfRows, true);
        }
        public IXLCells InsertCellsBelow(int numberOfRows, Boolean expandRange)
        {
            return InsertRowsBelow(numberOfRows, expandRange).Cells();
        }

        public Int32 CellCount()
        {
            return this.RangeAddress.LastAddress.ColumnNumber - this.RangeAddress.FirstAddress.ColumnNumber + 1;
        }
    }
}

