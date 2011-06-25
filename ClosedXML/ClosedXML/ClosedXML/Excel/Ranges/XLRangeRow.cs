namespace ClosedXML.Excel
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
            if (!quickLoad)
            {
                Worksheet.RangeShiftedRows += WorksheetRangeShiftedRows;
                Worksheet.RangeShiftedColumns += WorksheetRangeShiftedColumns;
                m_defaultStyle = new XLStyle(this, rangeParameters.DefaultStyle);
            }
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
            var retVal = new XLCells(false, false, false);
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
            return Cells(ExcelHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLRangeRow Sort()
        {
            AsRange().Sort(XLSortOrientation.LeftToRight);
            return this;
        }

        public IXLRangeRow Sort(XLSortOrder sortOrder)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder);
            return this;
        }

        public IXLRangeRow Sort(bool matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }

        public IXLRangeRow Sort(XLSortOrder sortOrder, bool matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }

        public new IXLRangeRow CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
                lastRowNumber = ExcelHelper.MaxRowNumber;
            int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
                lastColumnNumber = ExcelHelper.MaxColumnNumber;

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
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
                lastRowNumber = ExcelHelper.MaxRowNumber;
            int lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
                lastColumnNumber = ExcelHelper.MaxColumnNumber;

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

        public IXLRangeRows Rows(string rows)
        {
            var retVal = new XLRangeRows();
            var columnPairs = rows.Split(',');
            foreach (string pair in columnPairs)
            {
                string trimmedPair = pair.Trim();
                string firstColumn;
                string lastColumn;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    if (trimmedPair.Contains('-'))
                        trimmedPair = trimmedPair.Replace('-', ':');

                    var columnRange = trimmedPair.Split(':');
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
                bool thisCellIsBlank = StringExtensions.IsNullOrWhiteSpace(thisCell.InnerText);
                bool otherCellIsBlank = StringExtensions.IsNullOrWhiteSpace(otherCell.InnerText);
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
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                        else
                            comparison = Double.Parse(thisCell.InnerText).CompareTo(Double.Parse(otherCell.InnerText));
                    }
                    else if (e.MatchCase)
                        comparison = thisCell.GetString().ToLower().CompareTo(otherCell.GetString().ToLower());
                    else
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }

                if (comparison != 0)
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
            }

            return 0;
        }
    }
}