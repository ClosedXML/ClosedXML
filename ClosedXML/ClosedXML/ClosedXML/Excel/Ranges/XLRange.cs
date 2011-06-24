using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRange : XLRangeBase, IXLRange
    {
        #region Fields
        private IXLSortElements m_sortRows;
        private IXLSortElements m_sortColumns;
        #endregion
        #region Constructor
        public XLRange(XLRangeParameters xlRangeParameters)
                : base(xlRangeParameters.RangeAddress)
        {
            RangeParameters = xlRangeParameters;

            if (!xlRangeParameters.IgnoreEvents)
            {
                (Worksheet).RangeShiftedRows += Worksheet_RangeShiftedRows;
                (Worksheet).RangeShiftedColumns += Worksheet_RangeShiftedColumns;
                xlRangeParameters.IgnoreEvents = true;
            }
            m_defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
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
        #region IXLRange Members
        public IXLRangeColumn FirstColumn()
        {
            return Column(1);
        }
        public IXLRangeColumn LastColumn()
        {
            return Column(ColumnCount());
        }
        public IXLRangeColumn FirstColumnUsed()
        {
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var columnCount = ColumnCount();
            Int32 minColumnUsed = Int32.MaxValue;
            Int32 minColumnInCells = Int32.MaxValue;
            if ((Worksheet).Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
            {
                minColumnInCells = (Worksheet).Internals.CellsCollection
                        .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Min();
            }

            Int32 minCoInColumns = Int32.MaxValue;
            if ((Worksheet).Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
            {
                minCoInColumns = (Worksheet).Internals.ColumnsCollection
                        .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Min();
            }

            minColumnUsed = minColumnInCells < minCoInColumns ? minColumnInCells : minCoInColumns;

            if (minColumnUsed == Int32.MaxValue)
            {
                return null;
            }
            else
            {
                return Column(minColumnUsed);
            }
        }
        public IXLRangeColumn LastColumnUsed()
        {
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var columnCount = ColumnCount();
            Int32 maxColumnUsed = 0;
            Int32 maxColumnInCells = 0;
            if ((Worksheet).Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
            {
                maxColumnInCells = (Worksheet).Internals.CellsCollection
                        .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Max();
            }

            Int32 maxCoInColumns = 0;
            if ((Worksheet).Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
            {
                maxCoInColumns = (Worksheet).Internals.ColumnsCollection
                        .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Max();
            }

            maxColumnUsed = maxColumnInCells > maxCoInColumns ? maxColumnInCells : maxCoInColumns;

            if (maxColumnUsed == 0)
            {
                return null;
            }
            else
            {
                return Column(maxColumnUsed);
            }
        }

        public IXLRangeRow FirstRow()
        {
            return Row(1);
        }
        public IXLRangeRow LastRow()
        {
            return Row(RowCount());
        }
        public IXLRangeRow FirstRowUsed()
        {
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var rowCount = RowCount();
            Int32 minRowUsed = Int32.MaxValue;
            Int32 minRowInCells = Int32.MaxValue;
            if ((Worksheet).Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
            {
                minRowInCells = (Worksheet).Internals.CellsCollection
                        .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Min();
            }

            Int32 minRoInRows = Int32.MaxValue;
            if ((Worksheet).Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
            {
                minRoInRows = (Worksheet).Internals.RowsCollection
                        .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Min();
            }

            minRowUsed = minRowInCells < minRoInRows ? minRowInCells : minRoInRows;

            if (minRowUsed == Int32.MaxValue)
            {
                return null;
            }
            return Row(minRowUsed);
        }
        public IXLRangeRow LastRowUsed()
        {
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var rowCount = RowCount();
            Int32 maxRowUsed = 0;
            Int32 maxRowInCells = 0;
            if ((Worksheet).Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
            {
                maxRowInCells = (Worksheet).Internals.CellsCollection
                        .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Max();
            }

            Int32 maxRoInRows = 0;
            if ((Worksheet).Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
            {
                maxRoInRows = (Worksheet).Internals.RowsCollection
                        .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Max();
            }

            maxRowUsed = maxRowInCells > maxRoInRows ? maxRowInCells : maxRoInRows;

            if (maxRowUsed == 0)
            {
                return null;
            }
            else
            {
                return Row(maxRowUsed);
            }
        }

        public IXLRangeRow Row(Int32 row)
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
                    new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style));
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
        public IXLRangeColumn Column(Int32 column)
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
                    new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style));
        }
        public IXLRangeColumn Column(String column)
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

        public IXLRangeColumns Columns()
        {
            var retVal = new XLRangeColumns();
            foreach (var c in Enumerable.Range(1, ColumnCount()))
            {
                retVal.Add(Column(c));
            }
            return retVal;
        }
        public virtual IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRangeColumns();

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add(Column(co));
            }
            return retVal;
        }
        public IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return Columns(ExcelHelper.GetColumnNumberFromLetter(firstColumn), ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
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

                Int32 tmp;
                if (Int32.TryParse(firstColumn, out tmp))
                {
                    foreach (var col in Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                    {
                        retVal.Add(col);
                    }
                }
                else
                {
                    foreach (var col in Columns(firstColumn, lastColumn))
                    {
                        retVal.Add(col);
                    }
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
            foreach (var r in Enumerable.Range(1, RowCount()))
            {
                retVal.Add(Row(r));
            }
            return retVal;
        }
        public IXLRangeRows Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRangeRows();

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(Row(ro));
            }
            return retVal;
        }
        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
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
                foreach (var row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        public void Transpose(XLTransposeOptions transposeOption)
        {
            var rowCount = RowCount();
            var columnCount = ColumnCount();
            var squareSide = rowCount > columnCount ? rowCount : columnCount;

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

            foreach (var c in Range(1, 1, columnCount, rowCount).Cells())
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

        private void TransposeRange(int squareSide)
        {
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            XLRange rngToTranspose = (XLRange) Worksheet.Range(
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
                    var newKey = rngToTranspose.Cell(co, ro).Address; // new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
                    var newCell = new XLCell(Worksheet, newKey, oldCell.Style);
                    newCell.CopyFrom(oldCell);
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(oldCell.Address);
                }
            }
            //Int32 roInitial = rngToTranspose.RangeAddress.FirstAddress.RowNumber;
            //Int32 coInitial = rngToTranspose.RangeAddress.FirstAddress.ColumnNumber;
            //foreach (var c in rngToTranspose.Cells())
            //{
            //    var newKey = new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
            //    var newCell = new XLCell(newKey, c.Style, Worksheet);
            //    newCell.Value = c.Value;
            //    newCell.DataType = c.DataType;
            //    cellsToInsert.Add(newKey, newCell);
            //    cellsToDelete.Add(c.Address);
            //}
            cellsToDelete.ForEach(c => (Worksheet).Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => (Worksheet).Internals.CellsCollection.Add(c.Key, c.Value));
        }

        private void TransposeMerged(Int32 squareSide)
        {
            
            var rngToTranspose = new SheetRange(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                    RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            var mranges = new List<SheetRange>();
            foreach (var merge in (Worksheet).Internals.MergedRanges)
            {
                if (Contains(merge))
                {
                    mranges.Add(new SheetRange(merge.FirstAddress,
                                                           new SheetPoint(rngToTranspose.FirstAddress.RowNumber + merge.ColumnCount,
                                                                          rngToTranspose.FirstAddress.ColumnNumber + merge.RowCount)));

                }
            }
            mranges.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
            mranges.ForEach(m => Worksheet.Internals.MergedRanges.Add(m));
        }

        private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
        {
            if (transposeOption == XLTransposeOptions.MoveCells)
            {
                if (rowCount > columnCount)
                {
                    InsertColumnsAfter(rowCount - columnCount, false);
                }
                else if (columnCount > rowCount)
                {
                    InsertRowsBelow(columnCount - rowCount, false);
                }
            }
            else
            {
                if (rowCount > columnCount)
                {
                    var toMove = rowCount - columnCount;
                    var rngToClear = Worksheet.Range(
                            RangeAddress.FirstAddress.RowNumber,
                            RangeAddress.LastAddress.ColumnNumber + 1,
                            RangeAddress.LastAddress.RowNumber,
                            RangeAddress.LastAddress.ColumnNumber + toMove);
                    rngToClear.Clear();
                }
                else if (columnCount > rowCount)
                {
                    var toMove = columnCount - rowCount;
                    var rngToClear = Worksheet.Range(
                            RangeAddress.LastAddress.RowNumber + 1,
                            RangeAddress.FirstAddress.ColumnNumber,
                            RangeAddress.LastAddress.RowNumber + toMove,
                            RangeAddress.LastAddress.ColumnNumber);
                    rngToClear.Clear();
                }
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
        #endregion
        public override bool Equals(object obj)
        {
            var other = (XLRange) obj;
            return RangeAddress.Equals(other.RangeAddress)
                   && Worksheet.Equals(other.Worksheet);
        }

        public override int GetHashCode()
        {
            return RangeAddress.GetHashCode()
                   ^ Worksheet.GetHashCode();
        }

        public IXLSortElements SortRows
        {
            get { return m_sortRows ?? (m_sortRows = new XLSortElements()); }
        }

        public IXLSortElements SortColumns
        {
            get { return m_sortColumns ?? (m_sortColumns = new XLSortElements()); }
        }

        public IXLRange Sort()
        {
            if (SortColumns.Count() == 0)
            {
                return Sort(XLSortOrder.Ascending);
            }
            SortRangeRows();
            return this;
        }
        public IXLRange Sort(Boolean matchCase)
        {
            if (SortColumns.Count() == 0)
            {
                return Sort(XLSortOrder.Ascending, false);
            }
            SortRangeRows();
            return this;
        }
        public IXLRange Sort(XLSortOrder sortOrder)
        {
            if (SortColumns.Count() == 0)
            {
                Int32 columnCount = ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                {
                    SortColumns.Add(co, sortOrder);
                }
            }
            else
            {
                SortColumns.ForEach(sc => sc.SortOrder = sortOrder);
            }
            SortRangeRows();
            return this;
        }
        public IXLRange Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            if (SortColumns.Count() == 0)
            {
                Int32 columnCount = ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                {
                    SortColumns.Add(co, sortOrder, true, matchCase);
                }
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
            foreach (String coPair in columnsToSortBy.Split(','))
            {
                String coPairTrimmed = coPair.Trim();
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
                {
                    co = ExcelHelper.GetColumnNumberFromLetter(coString);
                }

                if (order.ToUpper().Equals("ASC"))
                {
                    SortColumns.Add(co, XLSortOrder.Ascending);
                }
                else
                {
                    SortColumns.Add(co, XLSortOrder.Descending);
                }
            }

            SortRangeRows();
            return this;
        }
        public IXLRange Sort(String columnsToSortBy, Boolean matchCase)
        {
            SortColumns.Clear();
            foreach (String coPair in columnsToSortBy.Split(','))
            {
                String coPairTrimmed = coPair.Trim();
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
                {
                    co = ExcelHelper.GetColumnNumberFromLetter(coString);
                }

                if (order.ToUpper().Equals("ASC"))
                {
                    SortColumns.Add(co, XLSortOrder.Ascending, true, matchCase);
                }
                else
                {
                    SortColumns.Add(co, XLSortOrder.Descending, true, matchCase);
                }
            }

            SortRangeRows();
            return this;
        }

        public IXLRange Sort(XLSortOrientation sortOrientation)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort();
            }
            if (SortRows.Count() == 0)
            {
                return Sort(sortOrientation, XLSortOrder.Ascending);
            }
            SortRangeColumns();
            return this;
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(matchCase);
            }
            if (SortRows.Count() == 0)
            {
                return Sort(sortOrientation, XLSortOrder.Ascending, matchCase);
            }
            SortRangeColumns();
            return this;
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(sortOrder);
            }
            if (SortRows.Count() == 0)
            {
                Int32 rowCount = RowCount();
                for (Int32 co = 1; co <= rowCount; co++)
                {
                    SortRows.Add(co, sortOrder);
                }
            }
            else
            {
                SortRows.ForEach(sc => sc.SortOrder = sortOrder);
            }
            SortRangeColumns();
            return this;
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(sortOrder, matchCase);
            }
            if (SortRows.Count() == 0)
            {
                Int32 rowCount = RowCount();
                for (Int32 co = 1; co <= rowCount; co++)
                {
                    SortRows.Add(co, sortOrder, matchCase);
                }
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
            {
                return Sort(elementsToSortBy);
            }
            SortRows.Clear();
            foreach (String roPair in elementsToSortBy.Split(','))
            {
                String roPairTrimmed = roPair.Trim();
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

                if (order.ToUpper().Equals("ASC"))
                {
                    SortRows.Add(ro, XLSortOrder.Ascending);
                }
                else
                {
                    SortRows.Add(ro, XLSortOrder.Descending);
                }
            }

            SortRangeColumns();
            return this;
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(elementsToSortBy, matchCase);
            }
            SortRows.Clear();
            foreach (String roPair in elementsToSortBy.Split(','))
            {
                String roPairTrimmed = roPair.Trim();
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

                if (order.ToUpper().Equals("ASC"))
                {
                    SortRows.Add(ro, XLSortOrder.Ascending, true, matchCase);
                }
                else
                {
                    SortRows.Add(ro, XLSortOrder.Descending, true, matchCase);
                }
            }

            SortRangeColumns();
            return this;
        }
        #region Sort Rows
        private void SortRangeRows()
        {
            SortingRangeRows(1, RowCount());
        }
        private void SwapRows(Int32 row1, Int32 row2)
        {
            Int32 cellCount = ColumnCount();

            for (Int32 co = 1; co <= cellCount; co++)
            {
                var cell1 = (XLCell) RowQuick(row1).Cell(co);
                var cell1Address = cell1.Address;
                var cell2 = (XLCell) RowQuick(row2).Cell(co);

                cell1.Address = cell2.Address;
                cell2.Address = cell1Address;

                (Worksheet).Internals.CellsCollection[cell1.Address] = cell1;
                (Worksheet).Internals.CellsCollection[cell2.Address] = cell2;
            }
        }
        private int SortRangeRows(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
            {
                m++;
            }

            while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
            {
                n--;
            }
            while (m < n)
            {
                SwapRows(m, n);

                while (m < endPoint && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                {
                    m++;
                }

                while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                {
                    n--;
                }
            }
            if (pivot != n)
            {
                SwapRows(n, pivot);
            }
            return n;
        }
        private void SortingRangeRows(int beg, int end)
        {
            if (end == beg)
            {
                return;
            }
            int pivot = SortRangeRows(beg, end);
            if (pivot > beg)
            {
                SortingRangeRows(beg, pivot - 1);
            }
            if (pivot < end)
            {
                SortingRangeRows(pivot + 1, end);
            }
        }
        #endregion
        #region Sort Columns
        private void SortRangeColumns()
        {
            SortingRangeColumns(1, ColumnCount());
        }
        private void SwapColumns(Int32 column1, Int32 column2)
        {
            Int32 cellCount = ColumnCount();

            for (Int32 co = 1; co <= cellCount; co++)
            {
                var cell1 = (XLCell) ColumnQuick(column1).Cell(co);
                var cell1Address = cell1.Address;
                var cell2 = (XLCell) ColumnQuick(column2).Cell(co);

                cell1.Address = cell2.Address;
                cell2.Address = cell1Address;

                (Worksheet).Internals.CellsCollection[cell1.Address] = cell1;
                (Worksheet).Internals.CellsCollection[cell2.Address] = cell2;
            }
        }
        private int SortRangeColumns(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && ColumnQuick(pivot).CompareTo((ColumnQuick(m)), SortRows) >= 0)
            {
                m++;
            }

            while ((n > begPoint) && ((ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0))
            {
                n--;
            }
            while (m < n)
            {
                SwapColumns(m, n);

                while ((m < endPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(m)), SortRows) >= 0)
                {
                    m++;
                }

                while ((n > begPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0)
                {
                    n--;
                }
            }
            if (pivot != n)
            {
                SwapColumns(n, pivot);
            }
            return n;
        }
        private void SortingRangeColumns(int beg, int end)
        {
            if (end == beg)
            {
                return;
            }
            int pivot = SortRangeColumns(beg, end);
            if (pivot > beg)
            {
                SortingRangeColumns(beg, pivot - 1);
            }
            if (pivot < end)
            {
                SortingRangeColumns(pivot + 1, end);
            }
        }
        #endregion
        public new IXLRange CopyTo(IXLCell target)
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
                                          lastColumnNumber);
        }
        public new IXLRange CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
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
                                          lastColumnNumber);
        }

        public IXLRange SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }
    }
}