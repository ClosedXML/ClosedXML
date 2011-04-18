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
        public XLRangeColumn(XLRangeParameters xlRangeParameters, Boolean quick)
            : base(xlRangeParameters.RangeAddress)
        {
            Worksheet = xlRangeParameters.Worksheet;
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

        public Int32 CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
        {
            foreach (var e in rowsToSort)
            {
                var thisCell = (XLCell)this.Cell(e.ElementNumber);
                var otherCell = (XLCell)otherColumn.Cell(e.ElementNumber);
                Int32 comparison;
                Boolean thisCellIsBlank = StringExtensions.IsNullOrWhiteSpace(thisCell.InnerText);
                Boolean otherCellIsBlank = StringExtensions.IsNullOrWhiteSpace(otherCell.InnerText);
                if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
                {
                    if (thisCellIsBlank && otherCellIsBlank)
                        comparison = 0;
                    else
                    {
                        if (thisCellIsBlank)
                            comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                        else
                            comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                    }
                }
                else
                {
                    if (thisCell.DataType == otherCell.DataType)
                    {
                        if (thisCell.DataType == XLCellValues.Text)
                        {
                            if (e.MatchCase)
                                comparison = thisCell.InnerText.CompareTo(otherCell.InnerText);
                            else
                                comparison = thisCell.InnerText.ToLower().CompareTo(otherCell.InnerText.ToLower());
                        }
                        else if (thisCell.DataType == XLCellValues.TimeSpan)
                        {
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                        }
                        else
                        {
                            comparison = Double.Parse(thisCell.InnerText).CompareTo(Double.Parse(otherCell.InnerText));
                        }
                    }
                    else
                        if (e.MatchCase)
                            comparison = thisCell.GetString().ToLower().CompareTo(otherCell.GetString().ToLower());
                        else
                            comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }
                if (comparison != 0)
                {
                    if (e.SortOrder == XLSortOrder.Ascending)
                        return comparison;
                    else
                        return comparison * -1;
                }
            }
            return 0;
        }

        public IXLRangeColumn Sort()
        {
            this.AsRange().Sort();
            return this;
        }
        public IXLRangeColumn Sort(XLSortOrder sortOrder)
        {
            this.AsRange().Sort(sortOrder);
            return this;
        }
        public IXLRangeColumn Sort(Boolean matchCase)
        {
            this.AsRange().Sort(matchCase);
            return this;
        }
        public IXLRangeColumn Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            this.AsRange().Sort(sortOrder, matchCase);
            return this;
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.Address.RowNumber + this.RowCount() - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows) lastRowNumber = XLWorksheet.MaxNumberOfRows;
            Int32 lastColumnNumber = target.Address.ColumnNumber + this.ColumnCount() - 1;
            if (lastColumnNumber > XLWorksheet.MaxNumberOfColumns) lastColumnNumber = XLWorksheet.MaxNumberOfColumns;

            return target.Worksheet.Range(target.Address.RowNumber, target.Address.ColumnNumber,
                lastRowNumber, lastColumnNumber)
                .Column(1);
        }
        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + this.RowCount() - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows) lastRowNumber = XLWorksheet.MaxNumberOfRows;
            Int32 lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + this.ColumnCount() - 1;
            if (lastColumnNumber > XLWorksheet.MaxNumberOfColumns) lastColumnNumber = XLWorksheet.MaxNumberOfColumns;

            return (target as XLRangeBase).Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.FirstAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Column(1);
        }
    }
}

