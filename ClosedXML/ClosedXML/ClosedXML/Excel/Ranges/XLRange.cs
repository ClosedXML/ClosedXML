using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRange : XLRangeBase, IXLRange
    {
        #region Fields

        private IXLSortElements _mSortColumns;
        private IXLSortElements _mSortRows;

        #endregion

        #region Constructor

        public XLRange(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters.RangeAddress)
        {
            RangeParameters = xlRangeParameters;

            if (!xlRangeParameters.IgnoreEvents)
            {
                Worksheet.RangeShiftedRows += WorksheetRangeShiftedRows;
                Worksheet.RangeShiftedColumns += WorksheetRangeShiftedColumns;
                xlRangeParameters.IgnoreEvents = true;
            }
            DefaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        #endregion

        public XLRangeParameters RangeParameters { get; private set; }

        #region IXLRange Members

        IXLRangeColumn IXLRange.FirstColumn()
        {
            return FirstColumn();
        }

        IXLRangeColumn IXLRange.LastColumn()
        {
            return LastColumn();
        }

        IXLRangeColumn IXLRange.FirstColumnUsed()
        {
            return FirstColumnUsed();
        }

        IXLRangeColumn IXLRange.FirstColumnUsed(bool includeFormats)
        {
            return FirstColumnUsed(includeFormats);
        }

        IXLRangeColumn IXLRange.LastColumnUsed()
        {
            return LastColumnUsed();
        }

        IXLRangeColumn IXLRange.LastColumnUsed(bool includeFormats)
        {
            return LastColumnUsed(includeFormats);
        }

        IXLRangeRow IXLRange.FirstRow()
        {
            return FirstRow();
        }

        IXLRangeRow IXLRange.LastRow()
        {
            return LastRow();
        }

        IXLRangeRow IXLRange.LastRowUsed()
        {
            return LastRowUsed();
        }

        IXLRangeRow IXLRange.LastRowUsed(bool includeFormats)
        {
            return LastRowUsed(includeFormats);
        }

        IXLRangeRow IXLRange.FirstRowUsed()
        {
            return FirstRowUsed();
        }

        IXLRangeRow IXLRange.FirstRowUsed(bool includeFormats)
        {
            return FirstRowUsed(includeFormats);
        }

        IXLRangeRow IXLRange.Row(Int32 row)
        {
            return Row(row);
        }

        IXLRangeColumn IXLRange.Column(Int32 column)
        {
            return Column(column);
        }

        IXLRangeColumn IXLRange.Column(String column)
        {
            return Column(column);
        }

        public IXLRangeColumns Columns()
        {
            var retVal = new XLRangeColumns();
            Int32 columnCount = ColumnCount();
            for (Int32 c = 1; c <= columnCount; c++ )
                retVal.Add(Column(c));
            return retVal;
        }

        public virtual IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRangeColumns();

            for (int co = firstColumn; co <= lastColumn; co++)
                retVal.Add(Column(co));
            return retVal;
        }

        public IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return Columns(ExcelHelper.GetColumnNumberFromLetter(firstColumn),
                           ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
            foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
            {
                String firstColumn;
                String lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    string[] columnRange = ExcelHelper.SplitRange(tPair);

                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
                }

                Int32 tmp;
                if (Int32.TryParse(firstColumn, out tmp))
                {
                    foreach (IXLRangeColumn col in Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                        retVal.Add(col);
                }
                else
                {
                    foreach (IXLRangeColumn col in Columns(firstColumn, lastColumn))
                        retVal.Add(col);
                }
            }
            return retVal;
        }

        IXLCell IXLRange.Cell(int row, int column)
        {
            return Cell(row, column);
        }

        IXLCell IXLRange.Cell(string cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLCell IXLRange.Cell(int row, string column)
        {
            return Cell(row, column);
        }

        IXLCell IXLRange.Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLRange IXLRange.Range(IXLRangeAddress rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLRange.Range(string rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLRange.Range(IXLCell firstCell, IXLCell lastCell)
        {
            return Range(firstCell, lastCell);
        }

        IXLRange IXLRange.Range(string firstCellAddress, string lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLRange.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLRange.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
        }

        public IXLRangeRows Rows()
        {
            var retVal = new XLRangeRows();
            Int32 rowCount = RowCount();
            for (Int32 r = 1; r <= rowCount; r++ )
                retVal.Add(Row(r));
            return retVal;
        }

        public IXLRangeRows Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRangeRows();

            for (int ro = firstRow; ro <= lastRow; ro++)
                retVal.Add(Row(ro));
            return retVal;
        }

        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
            {
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    string[] rowRange = ExcelHelper.SplitRange(tPair);

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (IXLRangeRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                    retVal.Add(row);
            }
            return retVal;
        }

        public void Transpose(XLTransposeOptions transposeOption)
        {
            int rowCount = RowCount();
            int columnCount = ColumnCount();
            int squareSide = rowCount > columnCount ? rowCount : columnCount;

            var firstCell = FirstCell();

            MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
            TransposeMerged(squareSide);
            TransposeRange(squareSide);
            RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                     firstCell.Address.RowNumber + columnCount - 1,
                                                     firstCell.Address.ColumnNumber + rowCount - 1,
                                                     RangeAddress.LastAddress.FixedRow,
                                                     RangeAddress.LastAddress.FixedColumn);
            if (rowCount > columnCount)
            {
                var rng = Worksheet.Range(
                    RangeAddress.LastAddress.RowNumber + 1,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                    RangeAddress.LastAddress.ColumnNumber);
                rng.Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
            else if (columnCount > rowCount)
            {
                var rng = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + 1,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
                rng.Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }

            foreach (IXLCell c in Range(1, 1, columnCount, rowCount).Cells())
            {
                var border = new XLBorder(this, c.Style.Border);
                c.Style.Border.TopBorder = border.LeftBorder;
                c.Style.Border.TopBorderColor = border.LeftBorderColor;
                c.Style.Border.LeftBorder = border.TopBorder;
                c.Style.Border.LeftBorderColor = border.TopBorderColor;
                c.Style.Border.RightBorder = border.BottomBorder;
                c.Style.Border.RightBorderColor = border.BottomBorderColor;
                c.Style.Border.BottomBorder = border.RightBorder;
                c.Style.Border.BottomBorderColor = border.RightBorderColor;
            }
        }

        public IXLTable AsTable()
        {
            return new XLTable(this, false);
        }

        public IXLTable AsTable(String name)
        {
            return new XLTable(this, name, false);
        }

        public IXLTable CreateTable()
        {
            return new XLTable(this, true);
        }

        public IXLTable CreateTable(String name)
        {
            return new XLTable(this, name, true);
        }

        public IXLSortElements SortRows
        {
            get { return _mSortRows ?? (_mSortRows = new XLSortElements()); }
        }

        public IXLSortElements SortColumns
        {
            get { return _mSortColumns ?? (_mSortColumns = new XLSortElements()); }
        }

        public IXLRange Sort()
        {
            if (SortColumns.Count() == 0)
                return Sort(XLSortOrder.Ascending);
            SortRangeRows();
            return this;
        }

        public IXLRange Sort(Boolean matchCase)
        {
            if (SortColumns.Count() == 0)
                return Sort(XLSortOrder.Ascending, false);
            SortRangeRows();
            return this;
        }

        public IXLRange Sort(XLSortOrder sortOrder)
        {
            if (SortColumns.Count() == 0)
            {
                Int32 columnCount = ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                    SortColumns.Add(co, sortOrder);
            }
            else
                SortColumns.ForEach(sc => sc.SortOrder = sortOrder);
            SortRangeRows();
            return this;
        }

        public IXLRange Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            if (SortColumns.Count() == 0)
            {
                Int32 columnCount = ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                    SortColumns.Add(co, sortOrder, true, matchCase);
            }
            else
            {
                SortColumns.ForEach(sc =>
                                        {
                                            sc.SortOrder = sortOrder;
                                            sc.MatchCase = matchCase;
                                        });
            }
            SortRangeRows();
            return this;
        }

        public IXLRange Sort(String columnsToSortBy)
        {
            SortColumns.Clear();
            foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                String coString;
                String order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = "ASC";
                }

                Int32 co;
                if (!Int32.TryParse(coString, out co))
                    co = ExcelHelper.GetColumnNumberFromLetter(coString);

                SortColumns.Add(co, String.Compare(order, "ASC", true) == 0 ? XLSortOrder.Ascending : XLSortOrder.Descending);
            }

            SortRangeRows();
            return this;
        }

        public IXLRange Sort(String columnsToSortBy, Boolean matchCase)
        {
            SortColumns.Clear();
            foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                String coString;
                String order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = "ASC";
                }

                Int32 co;
                if (!Int32.TryParse(coString, out co))
                    co = ExcelHelper.GetColumnNumberFromLetter(coString);

                SortColumns.Add(co, String.Compare(order, "ASC", true) == 0 ? XLSortOrder.Ascending : XLSortOrder.Descending, true,
                                matchCase);
            }

            SortRangeRows();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort();
            if (SortRows.Count() == 0)
                return Sort(sortOrientation, XLSortOrder.Ascending);
            SortRangeColumns();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort(matchCase);
            if (SortRows.Count() == 0)
                return Sort(sortOrientation, XLSortOrder.Ascending, matchCase);
            SortRangeColumns();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort(sortOrder);
            if (SortRows.Count() == 0)
            {
                Int32 rowCount = RowCount();
                for (Int32 co = 1; co <= rowCount; co++)
                    SortRows.Add(co, sortOrder);
            }
            else
                SortRows.ForEach(sc => sc.SortOrder = sortOrder);
            SortRangeColumns();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort(sortOrder, matchCase);
            if (SortRows.Count() == 0)
            {
                Int32 rowCount = RowCount();
                for (Int32 co = 1; co <= rowCount; co++)
                    SortRows.Add(co, sortOrder, matchCase);
            }
            else
            {
                SortRows.ForEach(sc =>
                                     {
                                         sc.SortOrder = sortOrder;
                                         sc.MatchCase = matchCase;
                                     });
            }
            SortRangeColumns();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort(elementsToSortBy);
            SortRows.Clear();
            foreach (string roPairTrimmed in elementsToSortBy.Split(',').Select(roPair => roPair.Trim()))
            {
                String roString;
                String order;
                if (roPairTrimmed.Contains(' '))
                {
                    var pair = roPairTrimmed.Split(' ');
                    roString = pair[0];
                    order = pair[1];
                }
                else
                {
                    roString = roPairTrimmed;
                    order = "ASC";
                }

                Int32 ro = Int32.Parse(roString);

                SortRows.Add(ro, String.Compare(order, "ASC",true)==0 ? XLSortOrder.Ascending : XLSortOrder.Descending);
            }

            SortRangeColumns();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
                return Sort(elementsToSortBy, matchCase);
            SortRows.Clear();
            foreach (string roPairTrimmed in elementsToSortBy.Split(',').Select(roPair => roPair.Trim()))
            {
                String roString;
                String order;
                if (roPairTrimmed.Contains(' '))
                {
                    var pair = roPairTrimmed.Split(' ');
                    roString = pair[0];
                    order = pair[1];
                }
                else
                {
                    roString = roPairTrimmed;
                    order = "ASC";
                }

                Int32 ro = Int32.Parse(roString);

                SortRows.Add(ro, String.Compare(order, "ASC", true) == 0 ? XLSortOrder.Ascending : XLSortOrder.Descending, true,
                             matchCase);
            }

            SortRangeColumns();
            return this;
        }

        public new IXLRange CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
                lastRowNumber = ExcelHelper.MaxRowNumber;
            Int32 lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
                lastColumnNumber = ExcelHelper.MaxColumnNumber;

            return target.Worksheet.Range(target.Address.RowNumber,
                                          target.Address.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber);
        }

        public new IXLRange CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > ExcelHelper.MaxRowNumber)
                lastRowNumber = ExcelHelper.MaxRowNumber;
            Int32 lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
                lastColumnNumber = ExcelHelper.MaxColumnNumber;

            return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                          target.RangeAddress.FirstAddress.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber);
        }

        public IXLRange SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }

        #endregion

        #region Sort Rows

        private void SortRangeRows()
        {
            SortingRangeRows(1, RowCount());
        }

        private void SwapRows(Int32 row1, Int32 row2)
        {
            int row1InWs = RangeAddress.FirstAddress.RowNumber + row1 - 1;
            int row2InWs = RangeAddress.FirstAddress.RowNumber + row2 - 1;

            Int32 firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumn = RangeAddress.LastAddress.ColumnNumber;

            var range1Sp1 = new XLSheetPoint(row1InWs, firstColumn);
            var range1Sp2 = new XLSheetPoint(row1InWs, lastColumn);
            var range2Sp1 = new XLSheetPoint(row2InWs, firstColumn);
            var range2Sp2 = new XLSheetPoint(row2InWs, lastColumn);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2));
        }

        private int SortRangeRows(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                m++;

            while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                n--;
            while (m < n)
            {
                SwapRows(m, n);

                while (m < endPoint && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                    m++;

                while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                    n--;
            }
            if (pivot != n)
                SwapRows(n, pivot);
            return n;
        }

        private void SortingRangeRows(int beg, int end)
        {
            if (end == beg)
                return;
            int pivot = SortRangeRows(beg, end);
            if (pivot > beg)
                SortingRangeRows(beg, pivot - 1);
            if (pivot < end)
                SortingRangeRows(pivot + 1, end);
        }

        #endregion

        #region Sort Columns

        private void SortRangeColumns()
        {
            SortingRangeColumns(1, ColumnCount());
        }

        private void SwapColumns(Int32 column1, Int32 column2)
        {
            int col1InWs = RangeAddress.FirstAddress.ColumnNumber + column1 - 1;
            int col2InWs = RangeAddress.FirstAddress.ColumnNumber + column2 - 1;

            Int32 firstRow = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRow = RangeAddress.LastAddress.RowNumber;

            var range1Sp1 = new XLSheetPoint(firstRow, col1InWs);
            var range1Sp2 = new XLSheetPoint(lastRow, col1InWs);
            var range2Sp1 = new XLSheetPoint(firstRow, col2InWs);
            var range2Sp2 = new XLSheetPoint(lastRow, col2InWs);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2));
        }

        private int SortRangeColumns(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && ColumnQuick(pivot).CompareTo((ColumnQuick(m)), SortRows) >= 0)
                m++;

            while ((n > begPoint) && ((ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0))
                n--;
            while (m < n)
            {
                SwapColumns(m, n);

                while ((m < endPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(m)), SortRows) >= 0)
                    m++;

                while ((n > begPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0)
                    n--;
            }
            if (pivot != n)
                SwapColumns(n, pivot);
            return n;
        }

        private void SortingRangeColumns(int beg, int end)
        {
            if (end == beg)
                return;
            int pivot = SortRangeColumns(beg, end);
            if (pivot > beg)
                SortingRangeColumns(beg, pivot - 1);
            if (pivot < end)
                SortingRangeColumns(pivot + 1, end);
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

        public XLRangeColumn FirstColumn()
        {
            return Column(1);
        }

        public IXLRangeColumn LastColumn()
        {
            return Column(ColumnCount());
        }

        public XLRangeColumn FirstColumnUsed()
        {
            return FirstColumnUsed(false);
        }

        public XLRangeColumn FirstColumnUsed(bool includeFormats)
        {
            var firstCellUsed = FirstCellUsed(includeFormats);
            return firstCellUsed == null ? null : Column(firstCellUsed.Address.ColumnNumber);
        }

        public XLRangeColumn LastColumnUsed()
        {
            return LastColumnUsed(false);
        }

        public XLRangeColumn LastColumnUsed(bool includeFormats)
        {
            var lastCellUsed = LastCellUsed(includeFormats);
            return lastCellUsed == null ? null : Column(lastCellUsed.Address.ColumnNumber);
        }

        public XLRangeRow FirstRow()
        {
            return Row(1);
        }

        public IXLRangeRow LastRow()
        {
            return Row(RowCount());
        }

        public XLRangeRow LastRowUsed()
        {
            return LastRowUsed(false);
        }

        public XLRangeRow LastRowUsed(bool includeFormats)
        {
            var lastCellUsed = LastCellUsed(includeFormats);
            return lastCellUsed == null ? null : Row(lastCellUsed.Address.RowNumber);
        }

        public XLRangeRow FirstRowUsed()
        {
            return FirstRowUsed(false);
        }

        public XLRangeRow FirstRowUsed(bool includeFormats)
        {
            var firstCellUsed = FirstCellUsed(includeFormats);
            return firstCellUsed == null ? null : Row(firstCellUsed.Address.RowNumber);
        }

        public XLRangeRow Row(Int32 row)
        {
            if (row <= 0 || row > ExcelHelper.MaxRowNumber)
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}", ExcelHelper.MaxRowNumber));

            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber + row - 1,
                                                 RangeAddress.FirstAddress.ColumnNumber,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.FirstAddress.RowNumber + row - 1,
                                                RangeAddress.LastAddress.ColumnNumber,
                                                false,
                                                false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), false);
        }

        public XLRangeRow RowQuick(Int32 row)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber + row - 1,
                                                 RangeAddress.FirstAddress.ColumnNumber,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.FirstAddress.RowNumber + row - 1,
                                                RangeAddress.LastAddress.ColumnNumber,
                                                false,
                                                false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), true);
        }

        public XLRangeColumn Column(Int32 column)
        {
            if (column <= 0 || column > ExcelHelper.MaxColumnNumber)
                throw new IndexOutOfRangeException(String.Format("Column number must be between 1 and {0}", ExcelHelper.MaxColumnNumber));

            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber,
                                                 RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.LastAddress.RowNumber,
                                                RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                false,
                                                false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), false);
        }

        public XLRangeColumn Column(String column)
        {
            return Column(ExcelHelper.GetColumnNumberFromLetter(column));
        }

        public XLRangeColumn ColumnQuick(Int32 column)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber,
                                                 RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.LastAddress.RowNumber,
                                                RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                false,
                                                false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), true);
        }

        private void TransposeRange(int squareSide)
        {
            var cellsToInsert = new Dictionary<XLSheetPoint, XLCell>();
            var cellsToDelete = new List<XLSheetPoint>();
            var rngToTranspose = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            Int32 roCount = rngToTranspose.RowCount();
            Int32 coCount = rngToTranspose.ColumnCount();
            for (Int32 ro = 1; ro <= roCount; ro++)
            {
                for (Int32 co = 1; co <= coCount; co++)
                {
                    var oldCell = rngToTranspose.Cell(ro, co);
                    var newKey = rngToTranspose.Cell(co, ro).Address;
                        // new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
                    var newCell = new XLCell(Worksheet, newKey, oldCell.Style);
                    newCell.CopyFrom(oldCell);
                    cellsToInsert.Add(new XLSheetPoint(newKey.RowNumber, newKey.ColumnNumber), newCell);
                    cellsToDelete.Add(new XLSheetPoint(oldCell.Address.RowNumber, oldCell.Address.ColumnNumber));
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
        }

        private void TransposeMerged(Int32 squareSide)
        {
            var rngToTranspose = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            foreach (IXLRange merge in Worksheet.Internals.MergedRanges.Where(Contains))
            {
                merge.RangeAddress.LastAddress = rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address;
            }
        }

        private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
        {
            if (transposeOption == XLTransposeOptions.MoveCells)
            {
                if (rowCount > columnCount)
                    InsertColumnsAfter(rowCount - columnCount, false);
                else if (columnCount > rowCount)
                    InsertRowsBelow(columnCount - rowCount, false);
            }
            else
            {
                if (rowCount > columnCount)
                {
                    int toMove = rowCount - columnCount;
                    var rngToClear = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber,
                        RangeAddress.LastAddress.ColumnNumber + 1,
                        RangeAddress.LastAddress.RowNumber,
                        RangeAddress.LastAddress.ColumnNumber + toMove);
                    rngToClear.Clear();
                }
                else if (columnCount > rowCount)
                {
                    int toMove = columnCount - rowCount;
                    var rngToClear = Worksheet.Range(
                        RangeAddress.LastAddress.RowNumber + 1,
                        RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber + toMove,
                        RangeAddress.LastAddress.ColumnNumber);
                    rngToClear.Clear();
                }
            }
        }

        public override bool Equals(object obj)
        {
            var other = (XLRange)obj;
            return RangeAddress.Equals(other.RangeAddress)
                   && Worksheet.Equals(other.Worksheet);
        }

        public override int GetHashCode()
        {
            return RangeAddress.GetHashCode()
                   ^ Worksheet.GetHashCode();
        }
    }
}