using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeRow : XLRangeBase, IXLRangeRow
    {
        #region Constructor
        public XLRangeRow(XLRangeParameters xlRangeParameters)
                : base(xlRangeParameters.RangeAddress)
        {
            RangeParameters = xlRangeParameters;

            (Worksheet).RangeShiftedRows += Worksheet_RangeShiftedRows;
            (Worksheet).RangeShiftedColumns += Worksheet_RangeShiftedColumns;
            m_defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }
        public XLRangeRow(XLRangeParameters xlRangeParameters, Boolean quick)
                : base(xlRangeParameters.RangeAddress)
        {
            RangeParameters = xlRangeParameters;
        }
        #endregion
        public XLRangeParameters RangeParameters { get; private set; }

        private void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(RangeAddress, range, columnsShifted);
        }
        private void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(RangeAddress, range, rowsShifted);
        }

        public IXLCell Cell(int column)
        {
            return Cell(1, column);
        }
        public new IXLCell Cell(string column)
        {
            return Cell(1, column);
        }

        public IXLRange Range(int firstColumn, int lastColumn)
        {
            return Range(1, firstColumn, 1, lastColumn);
        }

        public void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        public IXLCells InsertCellsAfter(int numberOfColumns)
        {
            return InsertCellsAfter(numberOfColumns, true);
        }
        public IXLCells InsertCellsAfter(int numberOfColumns, Boolean expandRange)
        {
            return InsertColumnsAfter(numberOfColumns, expandRange).Cells();
        }

        public IXLCells InsertCellsBefore(int numberOfColumns)
        {
            return InsertCellsBefore(numberOfColumns, false);
        }
        public IXLCells InsertCellsBefore(int numberOfColumns, Boolean expandRange)
        {
            return InsertColumnsBefore(numberOfColumns, expandRange).Cells();
        }

        public new IXLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, false, false);
            var rangePairs = cellsInRow.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }
            return retVal;
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
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixRowAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public IXLCells Cells(Int32 firstColumn, Int32 lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(String firstColumn, String lastColumn)
        {
            return Cells(ExcelHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public Int32 CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public Int32 CompareTo(XLRangeRow otherRow, IXLSortElements columnsToSort)
        {
            foreach (var e in columnsToSort)
            {
                var thisCell = (XLCell) Cell(e.ElementNumber);
                var otherCell = (XLCell) otherRow.Cell(e.ElementNumber);
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
        public IXLRangeRow Sort(Boolean matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }
        public IXLRangeRow Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }

        public new IXLRangeRow CopyTo(IXLCell target)
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
                    .Row(1);
        }
        public new IXLRangeRow CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);
            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
            {
                lastRowNumber = ExcelHelper.MaxRowNumber;
            }
            Int32 lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                lastColumnNumber = ExcelHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                          target.RangeAddress.LastAddress.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber)
                    .Row(1);
        }

        public IXLRangeRow Row(Int32 start, Int32 end)
        {
            return Range(1, start, 1, end).Row(1);
        }
        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var columnPairs = rows.Split(',');
            foreach (var pair in columnPairs)
            {
                var tPair = pair.Trim();
                String firstColumn;
                String lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                    {
                        tPair = tPair.Replace('-', ':');
                    }

                    var columnRange = tPair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
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
    }
}