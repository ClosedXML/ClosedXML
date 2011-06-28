namespace ClosedXML.Excel
{
    using System;
    using System.Linq;

    internal class XLRangeColumn : XLRangeBase, IXLRangeColumn
    {
        #region Constructor

        public XLRangeColumn(XLRangeParameters rangeParameters, bool quickLoad)
            : base(rangeParameters.RangeAddress)
        {
            if (!quickLoad)
            {
                Worksheet.RangeShiftedRows += WorksheetRangeShiftedRows;
                Worksheet.RangeShiftedColumns += WorksheetRangeShiftedColumns;
                DefaultStyle = new XLStyle(this, rangeParameters.DefaultStyle);
            }
        }

        #endregion

        #region IXLRangeColumn Members

        public XLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        IXLCell IXLRangeColumn.Cell(int row)
        {
            return Cell(row);
        }

        public new IXLCells Cells(string cellsInColumn)
        {
            var retVal = new XLCells( false, false);
            var rangePairs = cellsInColumn.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(int firstRow, int lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return InsertCellsAbove(numberOfRows, false);
        }

        public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange)
        {
            return InsertRowsAbove(numberOfRows, expandRange).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return InsertCellsBelow(numberOfRows, true);
        }

        public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange)
        {
            return InsertRowsBelow(numberOfRows, expandRange).Cells();
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLRangeColumn Sort()
        {
            AsRange().Sort();
            return this;
        }

        public IXLRangeColumn Sort(XLSortOrder sortOrder)
        {
            AsRange().Sort(sortOrder);
            return this;
        }

        public IXLRangeColumn Sort(bool matchCase)
        {
            AsRange().Sort(matchCase);
            return this;
        }

        public IXLRangeColumn Sort(XLSortOrder sortOrder, bool matchCase)
        {
            AsRange().Sort(sortOrder, matchCase);
            return this;
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
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
                .Column(1);
        }

        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
                lastRowNumber = ExcelHelper.MaxRowNumber;
            int lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
                lastColumnNumber = ExcelHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber, 
                target.RangeAddress.FirstAddress.ColumnNumber, 
                lastRowNumber, 
                lastColumnNumber)
                .Column(1);
        }

        public IXLRangeColumn Column(int start, int end)
        {
            return Range(start, end).FirstColumn();
        }

        public IXLRangeColumns Columns(string columns)
        {
            var retVal = new XLRangeColumns();
            var rowPairs = columns.Split(',');
            foreach (string pair in rowPairs)
            {
                string trimmedPair = pair.Trim();
                string firstRow;
                string lastRow;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    if (trimmedPair.Contains('-'))
                        trimmedPair = trimmedPair.Replace('-', ':');

                    var rowRange = trimmedPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = trimmedPair;
                    lastRow = trimmedPair;
                }

                retVal.Add(Range(firstRow, lastRow).FirstColumn());
            }

            return retVal;
        }

        public IXLRangeColumn SetDataType(XLCellValues dataType)
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

        public XLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
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
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public int CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
        {
            foreach (IXLSortElement e in rowsToSort)
            {
                var thisCell = Cell(e.ElementNumber);
                var otherCell = otherColumn.Cell(e.ElementNumber);
                int comparison;
                bool thisCellIsBlank = !thisCell.IsUsed();
                bool otherCellIsBlank = !otherCell.IsUsed();
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
                            comparison = e.MatchCase ? thisCell.InnerText.CompareTo(otherCell.InnerText) : thisCell.InnerText.ToLower().CompareTo(otherCell.InnerText.ToLower());
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