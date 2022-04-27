using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRange : XLRangeBase, IXLRange
    {
        #region Constructor

        public XLRange(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters.RangeAddress, (xlRangeParameters.DefaultStyle as XLStyle).Value)
        {
        }

        #endregion Constructor

        public override XLRangeType RangeType
        {
            get { return XLRangeType.Range; }
        }

        #region IXLRange Members

        IXLRangeRow IXLRange.Row(int row)
        {
            return Row(row);
        }

        IXLRangeColumn IXLRange.Column(int columnNumber)
        {
            return Column(columnNumber);
        }

        IXLRangeColumn IXLRange.Column(string columnLetter)
        {
            return Column(columnLetter);
        }

        public virtual IXLRangeColumns Columns(Func<IXLRangeColumn, bool> predicate = null)
        {
            var retVal = new XLRangeColumns();
            var columnCount = ColumnCount();
            for (var c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate == null || predicate(column))
                {
                    retVal.Add(column);
                }
            }
            return retVal;
        }

        public virtual IXLRangeColumns Columns(int firstColumn, int lastColumn)
        {
            var retVal = new XLRangeColumns();

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add(Column(co));
            }

            return retVal;
        }

        public virtual IXLRangeColumns Columns(string firstColumn, string lastColumn)
        {
            return Columns(XLHelper.GetColumnNumberFromLetter(firstColumn),
                           XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public virtual IXLRangeColumns Columns(string columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
            foreach (var tPair in columnPairs.Select(pair => pair.Trim()))
            {
                string firstColumn;
                string lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var columnRange = XLHelper.SplitRange(tPair);

                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
                }

                if (int.TryParse(firstColumn, out var tmp))
                {
                    foreach (var col in Columns(int.Parse(firstColumn), int.Parse(lastColumn)))
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

        public IXLRangeRows Rows(Func<IXLRangeRow, bool> predicate = null)
        {
            var retVal = new XLRangeRows();
            var rowCount = RowCount();
            for (var r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate == null || predicate(row))
                {
                    retVal.Add(Row(r));
                }
            }
            return retVal;
        }

        public IXLRangeRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLRangeRows();

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(Row(ro));
            }

            return retVal;
        }

        public IXLRangeRows Rows(string rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (var tPair in rowPairs.Select(pair => pair.Trim()))
            {
                string firstRow;
                string lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var rowRange = XLHelper.SplitRange(tPair);

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in Rows(int.Parse(firstRow), int.Parse(lastRow)))
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
            RangeAddress = new XLRangeAddress(
                RangeAddress.FirstAddress,
                new XLAddress(Worksheet,
                              firstCell.Address.RowNumber + columnCount - 1,
                              firstCell.Address.ColumnNumber + rowCount - 1,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));

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
                var border = (c.Style as XLStyle).Value.Border;
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
            return Worksheet.Table(this, false);
        }

        public IXLTable AsTable(string name)
        {
            return Worksheet.Table(this, name, false);
        }

        IXLTable IXLRange.CreateTable()
        {
            return CreateTable();
        }

        public XLTable CreateTable()
        {
            return (XLTable)Worksheet.Table(this, true, true);
        }

        IXLTable IXLRange.CreateTable(string name)
        {
            return CreateTable(name);
        }

        public XLTable CreateTable(string name)
        {
            return (XLTable)Worksheet.Table(this, name, true, true);
        }

        public IXLTable CreateTable(string name, bool setAutofilter)
        {
            return Worksheet.Table(this, name, true, setAutofilter);
        }

        public new IXLRange CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            var lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
            {
                lastRowNumber = XLHelper.MaxRowNumber;
            }

            var lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
            {
                lastColumnNumber = XLHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.Address.RowNumber,
                                          target.Address.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber);
        }

        public new IXLRange CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            var lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
            {
                lastRowNumber = XLHelper.MaxRowNumber;
            }

            var lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
            {
                lastColumnNumber = XLHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                          target.RangeAddress.FirstAddress.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber);
        }

        public IXLRange SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        public new IXLRange Sort()
        {
            return base.Sort().AsRange();
        }

        public new IXLRange Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            return base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
        }

        public new IXLRange Sort(int columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            return base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
        }

        public new IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            return base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks).AsRange();
        }

        #endregion IXLRange Members

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
        }

        IXLRangeColumn IXLRange.FirstColumn(Func<IXLRangeColumn, bool> predicate)
        {
            return FirstColumn(predicate);
        }

        internal XLRangeColumn FirstColumn(Func<IXLRangeColumn, bool> predicate = null)
        {
            if (predicate == null)
            {
                return Column(1);
            }

            var columnCount = ColumnCount();
            for (var c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate(column))
                {
                    return column;
                }
            }

            return null;
        }

        IXLRangeColumn IXLRange.LastColumn(Func<IXLRangeColumn, bool> predicate)
        {
            return LastColumn(predicate);
        }

        internal XLRangeColumn LastColumn(Func<IXLRangeColumn, bool> predicate = null)
        {
            var columnCount = ColumnCount();
            if (predicate == null)
            {
                return Column(columnCount);
            }

            for (var c = columnCount; c >= 1; c--)
            {
                var column = Column(c);
                if (predicate(column))
                {
                    return column;
                }
            }

            return null;
        }

        IXLRangeColumn IXLRange.FirstColumnUsed(Func<IXLRangeColumn, bool> predicate)
        {
            return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeColumn FirstColumnUsed(Func<IXLRangeColumn, bool> predicate = null)
        {
            return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn IXLRange.FirstColumnUsed(bool includeFormats, Func<IXLRangeColumn, bool> predicate)
        {
            return FirstColumnUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLRangeColumn IXLRange.FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate)
        {
            return FirstColumnUsed(options, predicate);
        }

        internal XLRangeColumn FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate = null)
        {
            if (predicate == null)
            {
                var firstColumnUsed = Worksheet.Internals.CellsCollection.FirstColumnUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return firstColumnUsed == 0 ? null : Column(firstColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
            }

            var columnCount = ColumnCount();
            for (var co = 1; co <= columnCount; co++)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && predicate(column))
                {
                    return column;
                }
            }
            return null;
        }

        IXLRangeColumn IXLRange.LastColumnUsed(Func<IXLRangeColumn, bool> predicate)
        {
            return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeColumn LastColumnUsed(Func<IXLRangeColumn, bool> predicate = null)
        {
            return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn IXLRange.LastColumnUsed(bool includeFormats, Func<IXLRangeColumn, bool> predicate)
        {
            return LastColumnUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLRangeColumn IXLRange.LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate)
        {
            return LastColumnUsed(options, predicate);
        }

        internal XLRangeColumn LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate = null)
        {
            if (predicate == null)
            {
                var lastColumnUsed = Worksheet.Internals.CellsCollection.LastColumnUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return lastColumnUsed == 0 ? null : Column(lastColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
            }

            var columnCount = ColumnCount();
            for (var co = columnCount; co >= 1; co--)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && predicate(column))
                {
                    return column;
                }
            }
            return null;
        }

        IXLRangeRow IXLRange.FirstRow(Func<IXLRangeRow, bool> predicate)
        {
            return FirstRow(predicate);
        }

        public XLRangeRow FirstRow(Func<IXLRangeRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                return Row(1);
            }

            var rowCount = RowCount();
            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);
                if (predicate(row))
                {
                    return row;
                }
            }

            return null;
        }

        IXLRangeRow IXLRange.LastRow(Func<IXLRangeRow, bool> predicate)
        {
            return LastRow(predicate);
        }

        public XLRangeRow LastRow(Func<IXLRangeRow, bool> predicate = null)
        {
            var rowCount = RowCount();
            if (predicate == null)
            {
                return Row(rowCount);
            }

            for (var ro = rowCount; ro >= 1; ro--)
            {
                var row = Row(ro);
                if (predicate(row))
                {
                    return row;
                }
            }

            return null;
        }

        IXLRangeRow IXLRange.FirstRowUsed(Func<IXLRangeRow, bool> predicate)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeRow FirstRowUsed(Func<IXLRangeRow, bool> predicate = null)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow IXLRange.FirstRowUsed(bool includeFormats, Func<IXLRangeRow, bool> predicate)
        {
            return FirstRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRow IXLRange.FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate)
        {
            return FirstRowUsed(options, predicate);
        }

        internal XLRangeRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                var rowFromCells = Worksheet.Internals.CellsCollection.FirstRowUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                //Int32 rowFromRows = Worksheet.Internals.RowsCollection.FirstRowUsed(includeFormats);

                return rowFromCells == 0 ? null : Row(rowFromCells - RangeAddress.FirstAddress.RowNumber + 1);
            }

            var rowCount = RowCount();
            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && predicate(row))
                {
                    return row;
                }
            }
            return null;
        }

        IXLRangeRow IXLRange.LastRowUsed(Func<IXLRangeRow, bool> predicate)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeRow LastRowUsed(Func<IXLRangeRow, bool> predicate = null)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow IXLRange.LastRowUsed(bool includeFormats, Func<IXLRangeRow, bool> predicate)
        {
            return LastRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRow IXLRange.LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate)
        {
            return LastRowUsed(options, predicate);
        }

        internal XLRangeRow LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                var lastRowUsed = Worksheet.Internals.CellsCollection.LastRowUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return lastRowUsed == 0 ? null : Row(lastRowUsed - RangeAddress.FirstAddress.RowNumber + 1);
            }

            var rowCount = RowCount();
            for (var ro = rowCount; ro >= 1; ro--)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && predicate(row))
                {
                    return row;
                }
            }
            return null;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRows IXLRange.RowsUsed(bool includeFormats, Func<IXLRangeRow, bool> predicate)
        {
            return RowsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRows IXLRange.RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate)
        {
            return RowsUsed(options, predicate);
        }

        internal XLRangeRows RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool> predicate = null)
        {
            var rows = new XLRangeRows();
            var rowCount = RowCount(options);

            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                {
                    rows.Add(row);
                }
            }
            return rows;
        }

        IXLRangeRows IXLRange.RowsUsed(Func<IXLRangeRow, bool> predicate)
        {
            return RowsUsed(predicate);
        }

        internal XLRangeRows RowsUsed(Func<IXLRangeRow, bool> predicate = null)
        {
            return RowsUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeColumns IXLRange.ColumnsUsed(bool includeFormats, Func<IXLRangeColumn, bool> predicate)
        {
            return ColumnsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeColumns IXLRange.ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate)
        {
            return ColumnsUsed(options, predicate);
        }

        internal virtual XLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = new XLRangeColumns();
            var columnCount = ColumnCount(options);

            for (var co = 1; co <= columnCount; co++)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                {
                    columns.Add(column);
                }
            }
            return columns;
        }

        IXLRangeColumns IXLRange.ColumnsUsed(Func<IXLRangeColumn, bool> predicate)
        {
            return ColumnsUsed(predicate);
        }

        internal virtual XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool> predicate = null)
        {
            return ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        public XLRangeRow Row(int row)
        {
            if (row <= 0 || row > XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            {
                throw new ArgumentOutOfRangeException(nameof(row), string.Format("Row number must be between 1 and {0}", XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1));
            }

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
            return Worksheet.RangeRow(new XLRangeAddress(firstCellAddress, lastCellAddress));
        }

        public virtual XLRangeColumn Column(int columnNumber)
        {
            if (columnNumber <= 0 || columnNumber > XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber), string.Format("Column number must be between 1 and {0}", XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1));
            }

            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber,
                                                 RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.LastAddress.RowNumber,
                                                RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
                                                false,
                                                false);
            return Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
        }

        public virtual XLRangeColumn Column(string columnLetter)
        {
            return Column(XLHelper.GetColumnNumberFromLetter(columnLetter));
        }

        internal IEnumerable<XLRange> Split(IXLRangeAddress anotherRange, bool includeIntersection)
        {
            if (!RangeAddress.Intersects(anotherRange))
            {
                yield return this;
                yield break;
            }

            var thisRow1 = RangeAddress.FirstAddress.RowNumber;
            var thisRow2 = RangeAddress.LastAddress.RowNumber;
            var thisColumn1 = RangeAddress.FirstAddress.ColumnNumber;
            var thisColumn2 = RangeAddress.LastAddress.ColumnNumber;

            var otherRow1 = Math.Min(Math.Max(thisRow1, anotherRange.FirstAddress.RowNumber), thisRow2 + 1);
            var otherRow2 = Math.Max(Math.Min(thisRow2, anotherRange.LastAddress.RowNumber), thisRow1 - 1);
            var otherColumn1 = Math.Min(Math.Max(thisColumn1, anotherRange.FirstAddress.ColumnNumber), thisColumn2 + 1);
            var otherColumn2 = Math.Max(Math.Min(thisColumn2, anotherRange.LastAddress.ColumnNumber), thisColumn1 - 1);

            var candidates = new[]
            {
                // to the top of the intersection
                new XLRangeAddress(
                    new XLAddress(thisRow1,thisColumn1, false, false),
                    new XLAddress(otherRow1 - 1, thisColumn2, false, false)),

                // to the left of the intersection
                new XLRangeAddress(
                    new XLAddress(otherRow1,thisColumn1, false, false),
                    new XLAddress(otherRow2, otherColumn1 - 1, false, false)),

                includeIntersection
                    ? new XLRangeAddress(
                        new XLAddress(otherRow1, otherColumn1, false, false),
                        new XLAddress(otherRow2, otherColumn2, false, false))
                    : XLRangeAddress.Invalid,

                // to the right of the intersection
                new XLRangeAddress(
                    new XLAddress(otherRow1,otherColumn2 + 1, false, false),
                    new XLAddress(otherRow2, thisColumn2, false, false)),

                // to the bottom of the intersection
                new XLRangeAddress(
                    new XLAddress(otherRow2 + 1,thisColumn1, false, false),
                    new XLAddress(thisRow2, thisColumn2, false, false)),
            };

            foreach (var rangeAddress in candidates.Where(c => c.IsValid && c.IsNormalized))
            {
                yield return Worksheet.Range(rangeAddress);
            }
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

            var roCount = rngToTranspose.RowCount();
            var coCount = rngToTranspose.ColumnCount();
            for (var ro = 1; ro <= roCount; ro++)
            {
                for (var co = 1; co <= coCount; co++)
                {
                    var oldCell = rngToTranspose.Cell(ro, co);
                    var newKey = rngToTranspose.Cell(co, ro).Address;
                    // new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
                    var newCell = new XLCell(Worksheet, newKey, oldCell.StyleValue);
                    newCell.CopyFrom(oldCell, XLCellCopyOptions.All);
                    cellsToInsert.Add(new XLSheetPoint(newKey.RowNumber, newKey.ColumnNumber), newCell);
                    cellsToDelete.Add(new XLSheetPoint(oldCell.Address.RowNumber, oldCell.Address.ColumnNumber));
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
        }

        private void TransposeMerged(int squareSide)
        {
            var rngToTranspose = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            foreach (var merge in Worksheet.Internals.MergedRanges.Where(Contains).Cast<XLRange>())
            {
                merge.RangeAddress = new XLRangeAddress(
                    merge.RangeAddress.FirstAddress,
                    rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address);
            }
        }

        private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
        {
            if (transposeOption == XLTransposeOptions.MoveCells)
            {
                if (rowCount > columnCount)
                {
                    InsertColumnsAfter(false, rowCount - columnCount, false);
                }
                else if (columnCount > rowCount)
                {
                    InsertRowsBelow(false, columnCount - rowCount, false);
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

        public override bool Equals(object obj)
        {
            if (!(obj is XLRange other))
            {
                return false;
            }

            return RangeAddress.Equals(other.RangeAddress)
                   && Worksheet.Equals(other.Worksheet);
        }

        public override int GetHashCode()
        {
            return RangeAddress.GetHashCode()
                   ^ Worksheet.GetHashCode();
        }

        public new IXLRange Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLRangeColumn FindColumn(Func<IXLRangeColumn, bool> predicate)
        {
            var columnCount = ColumnCount();
            for (var c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate == null || predicate(column))
                {
                    return column;
                }
            }
            return null;
        }

        public IXLRangeRow FindRow(Func<IXLRangeRow, bool> predicate)
        {
            var rowCount = RowCount();
            for (var r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate(row))
                {
                    return row;
                }
            }
            return null;
        }

        public override string ToString()
        {
            if (IsEntireSheet())
            {
                return Worksheet.Name;
            }
            else if (IsEntireRow())
            {
                return string.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    RangeAddress.FirstAddress.RowNumber,
                    ':',
                    RangeAddress.LastAddress.RowNumber);
            }
            else if (IsEntireColumn())
            {
                return string.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    RangeAddress.FirstAddress.ColumnLetter,
                    ':',
                    RangeAddress.LastAddress.ColumnLetter);
            }
            else
            {
                return base.ToString();
            }
        }
    }
}
