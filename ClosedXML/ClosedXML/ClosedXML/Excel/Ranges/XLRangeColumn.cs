using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumn : XLRangeBase, IXLRangeColumn
    {
        #region Constructor
        public XLRangeColumn(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters.RangeAddress)
        {
            (Worksheet).RangeShiftedRows += Worksheet_RangeShiftedRows;
            (Worksheet).RangeShiftedColumns += Worksheet_RangeShiftedColumns;
            m_defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }
        public XLRangeColumn(XLRangeParameters xlRangeParameters, Boolean quick)
            : base(xlRangeParameters.RangeAddress)
        {
        }
        #endregion

        private void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(RangeAddress, range, columnsShifted);
        }
        private void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(RangeAddress, range, rowsShifted);
        }

        public IXLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        public new IXLCells Cells(String cellsInColumn)
        {
            var retVal = new XLCells(false, false, false);
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

        public XLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
        }
        public override XLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                {
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');
                }

                String[] arrRange = rangeAddressStr.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
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
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public Int32 CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
        {
            foreach (var e in rowsToSort)
            {
                var thisCell = (XLCell) Cell(e.ElementNumber);
                var otherCell = (XLCell) otherColumn.Cell(e.ElementNumber);
                Int32 comparison;
                Boolean thisCellIsBlank = StringExtensions.IsNullOrWhiteSpace(thisCell.InnerText);
                Boolean otherCellIsBlank = StringExtensions.IsNullOrWhiteSpace(otherCell.InnerText);
                if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
                {
                    if (thisCellIsBlank && otherCellIsBlank)
                    {
                        comparison = 0;
                    }
                    else
                    {
                        if (thisCellIsBlank)
                        {
                            comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                        }
                        else
                        {
                            comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                        }
                    }
                }
                else
                {
                    if (thisCell.DataType == otherCell.DataType)
                    {
                        if (thisCell.DataType == XLCellValues.Text)
                        {
                            if (e.MatchCase)
                            {
                                comparison = thisCell.InnerText.CompareTo(otherCell.InnerText);
                            }
                            else
                            {
                                comparison = thisCell.InnerText.ToLower().CompareTo(otherCell.InnerText.ToLower());
                            }
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
                    else if (e.MatchCase)
                    {
                        comparison = thisCell.GetString().ToLower().CompareTo(otherCell.GetString().ToLower());
                    }
                    else
                    {
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                    }
                }
                if (comparison != 0)
                {
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison*-1;
                }
            }
            return 0;
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
        public IXLRangeColumn Sort(Boolean matchCase)
        {
            AsRange().Sort(matchCase);
            return this;
        }
        public IXLRangeColumn Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            AsRange().Sort(sortOrder, matchCase);
            return this;
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
            {
                lastRowNumber = ExcelHelper.MaxRowNumber;
            }
            Int32 lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                lastColumnNumber = ExcelHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.Address.RowNumber,
                                          target.Address.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber)
                    .Column(1);
        }
        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            var lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
            {
                lastRowNumber = ExcelHelper.MaxRowNumber;
            }
            Int32 lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                lastColumnNumber = ExcelHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                          target.RangeAddress.FirstAddress.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber)
                    .Column(1);
        }

        public IXLRangeColumn Column(Int32 start, Int32 end)
        {
            return Range(start, end).FirstColumn();
        }
        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var rowPairs = columns.Split(',');
            foreach (var pair in rowPairs)
            {
                var tPair = pair.Trim();
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                    {
                        tPair = tPair.Replace('-', ':');
                    }

                    var rowRange = tPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
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
    }
}