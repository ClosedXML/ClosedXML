﻿namespace ClosedXML.Excel
{
    using System;
    using System.Linq;
	

    internal class XLRangeRow : XLRangeBase, IXLRangeRow
    {
        #region Constructor

        public XLRangeRow(XLRangeParameters rangeParameters, bool quickLoad)
            : base(rangeParameters.RangeAddress)
        {
            RangeParameters = rangeParameters;
            if (quickLoad) return;
            if (!RangeParameters.IgnoreEvents)
            {
                SubscribeToShiftedRows((range, rowsShifted) => this.WorksheetRangeShiftedRows(range, rowsShifted));
                SubscribeToShiftedColumns((range, columnsShifted) => this.WorksheetRangeShiftedColumns(range, columnsShifted));
            }
            SetStyle(rangeParameters.DefaultStyle);
        }

        #endregion

        
        public XLRangeParameters RangeParameters { get; private set; }

        #region IXLRangeRow Members

        public IXLCell Cell(int column)
        {
            return Cell(1, column);
        }

        public new IXLCell Cell(string column)
        {
            return Cell(1, column);
        }

        public void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        public IXLCells InsertCellsAfter(int numberOfColumns)
        {
            return InsertCellsAfter(numberOfColumns, true);
        }

        public IXLCells InsertCellsAfter(int numberOfColumns, bool expandRange)
        {
            return InsertColumnsAfter(numberOfColumns, expandRange).Cells();
        }

        public IXLCells InsertCellsBefore(int numberOfColumns)
        {
            return InsertCellsBefore(numberOfColumns, false);
        }

        public IXLCells InsertCellsBefore(int numberOfColumns, bool expandRange)
        {
            return InsertColumnsBefore(numberOfColumns, expandRange).Cells();
        }

        public new IXLCells Cells(string cellsInRow)
        {
            var retVal = new XLCells(false, false);
            var rangePairs = cellsInRow.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(int firstColumn, int lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(string firstColumn, string lastColumn)
        {
            return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public new IXLRangeRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLRangeRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        public new IXLRangeRow CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Row(1);
        }

        public new IXLRangeRow CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);
            int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.LastAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Row(1);
        }

        public IXLRangeRow Row(int start, int end)
        {
            return Range(1, start, 1, end).Row(1);
        }

        public IXLRangeRow Row(IXLCell start, IXLCell end)
        {
            return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
        }

        public IXLRangeRows Rows(string rows)
        {
            var retVal = new XLRangeRows();
            var columnPairs = rows.Split(',');
            foreach (string trimmedPair in columnPairs.Select(pair => pair.Trim()))
            {
                string firstColumn;
                string lastColumn;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    var columnRange = trimmedPair.Contains('-')
                                          ? trimmedPair.Replace('-', ':').Split(':')
                                          : trimmedPair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = trimmedPair;
                    lastColumn = trimmedPair;
                }

                retVal.Add(Range(firstColumn, lastColumn).FirstRow());
            }

            return retVal;
        }

        public IXLRangeRow SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }

        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(RangeAddress.FirstAddress.RowNumber);
        }

        #endregion

        private void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(RangeAddress, range, columnsShifted);
        }

        private void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(RangeAddress, range, rowsShifted);
        }

        public IXLRange Range(int firstColumn, int lastColumn)
        {
            return Range(1, firstColumn, 1, lastColumn);
        }

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                var arrRange = rangeAddressStr.Split(':');
                string firstPart = arrRange[0];
                string secondPart = arrRange[1];
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
                rangeAddressToUse = FixRowAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public int CompareTo(XLRangeRow otherRow, IXLSortElements columnsToSort)
        {
            foreach (IXLSortElement e in columnsToSort)
            {
                var thisCell = (XLCell)Cell(e.ElementNumber);
                var otherCell = (XLCell)otherRow.Cell(e.ElementNumber);
                int comparison;
                bool thisCellIsBlank = thisCell.IsEmpty();
                bool otherCellIsBlank = otherCell.IsEmpty();
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
                            comparison = e.MatchCase
                                             ? thisCell.InnerText.CompareTo(otherCell.InnerText)
                                             : String.Compare(thisCell.InnerText, otherCell.InnerText, true);
                        }
                        else if (thisCell.DataType == XLCellValues.TimeSpan)
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                        else
                            comparison = Double.Parse(thisCell.InnerText).CompareTo(Double.Parse(otherCell.InnerText));
                    }
                    else if (e.MatchCase)
                        comparison = String.Compare(thisCell.GetString(), otherCell.GetString(), true);
                    else
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }

                if (comparison != 0)
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
            }

            return 0;
        }

        private XLRangeRow RowShift(Int32 rowsToShift)
        {
            Int32 rowNum = RowNumber() + rowsToShift;

            var range = Worksheet.Range(
                rowNum,
                RangeAddress.FirstAddress.ColumnNumber,
                rowNum,
                RangeAddress.LastAddress.ColumnNumber);
                
            var result = range.FirstRow();
            range.Dispose();

            return result;
        }

        #region XLRangeRow Above

        IXLRangeRow IXLRangeRow.RowAbove()
        {
            return RowAbove();
        }

        IXLRangeRow IXLRangeRow.RowAbove(Int32 step)
        {
            return RowAbove(step);
        }

        public XLRangeRow RowAbove()
        {
            return RowAbove(1);
        }

        public XLRangeRow RowAbove(Int32 step)
        {
            return RowShift(step * -1);
        }

        #endregion

        #region XLRangeRow Below

        IXLRangeRow IXLRangeRow.RowBelow()
        {
            return RowBelow();
        }

        IXLRangeRow IXLRangeRow.RowBelow(Int32 step)
        {
            return RowBelow(step);
        }

        public XLRangeRow RowBelow()
        {
            return RowBelow(1);
        }

        public XLRangeRow RowBelow(Int32 step)
        {
            return RowShift(step);
        }

        #endregion

        public new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLRangeRow RowUsed(Boolean includeFormats = false)
        {
            return Row(FirstCellUsed(includeFormats), LastCellUsed(includeFormats));
        }

    }
}