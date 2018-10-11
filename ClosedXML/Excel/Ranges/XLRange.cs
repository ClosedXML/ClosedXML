using ClosedXML.Extensions;
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

        IXLRangeRow IXLRange.Row(Int32 row)
        {
            return Row(row);
        }

        IXLRangeColumn IXLRange.Column(Int32 columnNumber)
        {
            return Column(columnNumber);
        }

        IXLRangeColumn IXLRange.Column(String columnLetter)
        {
            return Column(columnLetter);
        }

        public virtual IXLRangeColumns Columns(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            var retVal = new XLRangeColumns();
            Int32 columnCount = ColumnCount();
            for (Int32 c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate == null || predicate(column))
                    retVal.Add(column);
            }
            return retVal;
        }

        public virtual IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRangeColumns();

            for (int co = firstColumn; co <= lastColumn; co++)
                retVal.Add(Column(co));
            return retVal;
        }

        public virtual IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return Columns(XLHelper.GetColumnNumberFromLetter(firstColumn),
                           XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public virtual IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
            foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
            {
                String firstColumn;
                String lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    string[] columnRange = XLHelper.SplitRange(tPair);

                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
                }

                if (Int32.TryParse(firstColumn, out Int32 tmp))
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

        public IXLRangeRows Rows(Func<IXLRangeRow, Boolean> predicate = null)
        {
            var retVal = new XLRangeRows();
            Int32 rowCount = RowCount();
            for (Int32 r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate == null || predicate(row))
                    retVal.Add(Row(r));
            }
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
                    string[] rowRange = XLHelper.SplitRange(tPair);

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

            foreach (IXLCell c in Range(1, 1, columnCount, rowCount).Cells())
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

        public IXLTable AsTable(String name)
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

        IXLTable IXLRange.CreateTable(String name)
        {
            return CreateTable(name);
        }

        public XLTable CreateTable(String name)
        {
            return (XLTable)Worksheet.Table(this, name, true, true);
        }

        public IXLTable CreateTable(String name, Boolean setAutofilter)
        {
            return Worksheet.Table(this, name, true, setAutofilter);
        }

        public new IXLRange CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            Int32 lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(target.Address.RowNumber,
                                          target.Address.ColumnNumber,
                                          lastRowNumber,
                                          lastColumnNumber);
        }

        public new IXLRange CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            Int32 lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

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

        public new IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            return base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
        }

        public new IXLRange Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            return base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
        }

        public new IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
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

        IXLRangeColumn IXLRange.FirstColumn(Func<IXLRangeColumn, Boolean> predicate)
        {
            return FirstColumn(predicate);
        }

        internal XLRangeColumn FirstColumn(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            if (predicate == null)
                return Column(1);

            Int32 columnCount = ColumnCount();
            for (Int32 c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate(column)) return column;
            }

            return null;
        }

        IXLRangeColumn IXLRange.LastColumn(Func<IXLRangeColumn, Boolean> predicate)
        {
            return LastColumn(predicate);
        }

        internal XLRangeColumn LastColumn(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            Int32 columnCount = ColumnCount();
            if (predicate == null)
                return Column(columnCount);

            for (Int32 c = columnCount; c >= 1; c--)
            {
                var column = Column(c);
                if (predicate(column)) return column;
            }

            return null;
        }

        IXLRangeColumn IXLRange.FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate)
        {
            return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeColumn FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn IXLRange.FirstColumnUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate)
        {
            return FirstColumnUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLRangeColumn IXLRange.FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate)
        {
            return FirstColumnUsed(options, predicate);
        }

        internal XLRangeColumn FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null)
        {
            if (predicate == null)
            {
                Int32 firstColumnUsed = Worksheet.Internals.CellsCollection.FirstColumnUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return firstColumnUsed == 0 ? null : Column(firstColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
            }

            Int32 columnCount = ColumnCount();
            for (Int32 co = 1; co <= columnCount; co++)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && predicate(column))
                    return column;
            }
            return null;
        }

        IXLRangeColumn IXLRange.LastColumnUsed(Func<IXLRangeColumn, Boolean> predicate)
        {
            return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeColumn LastColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn IXLRange.LastColumnUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate)
        {
            return LastColumnUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLRangeColumn IXLRange.LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate)
        {
            return LastColumnUsed(options, predicate);
        }

        internal XLRangeColumn LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null)
        {
            if (predicate == null)
            {
                Int32 lastColumnUsed = Worksheet.Internals.CellsCollection.LastColumnUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return lastColumnUsed == 0 ? null : Column(lastColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
            }

            Int32 columnCount = ColumnCount();
            for (Int32 co = columnCount; co >= 1; co--)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && predicate(column))
                    return column;
            }
            return null;
        }

        IXLRangeRow IXLRange.FirstRow(Func<IXLRangeRow, Boolean> predicate)
        {
            return FirstRow(predicate);
        }

        public XLRangeRow FirstRow(Func<IXLRangeRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return Row(1);

            Int32 rowCount = RowCount();
            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);
                if (predicate(row)) return row;
            }

            return null;
        }

        IXLRangeRow IXLRange.LastRow(Func<IXLRangeRow, Boolean> predicate)
        {
            return LastRow(predicate);
        }

        public XLRangeRow LastRow(Func<IXLRangeRow, Boolean> predicate = null)
        {
            Int32 rowCount = RowCount();
            if (predicate == null)
                return Row(rowCount);

            for (Int32 ro = rowCount; ro >= 1; ro--)
            {
                var row = Row(ro);
                if (predicate(row)) return row;
            }

            return null;
        }

        IXLRangeRow IXLRange.FirstRowUsed(Func<IXLRangeRow, Boolean> predicate)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeRow FirstRowUsed(Func<IXLRangeRow, Boolean> predicate = null)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow IXLRange.FirstRowUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate)
        {
            return FirstRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRow IXLRange.FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate)
        {
            return FirstRowUsed(options, predicate);
        }

        internal XLRangeRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null)
        {
            if (predicate == null)
            {
                Int32 rowFromCells = Worksheet.Internals.CellsCollection.FirstRowUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                //Int32 rowFromRows = Worksheet.Internals.RowsCollection.FirstRowUsed(includeFormats);

                return rowFromCells == 0 ? null : Row(rowFromCells - RangeAddress.FirstAddress.RowNumber + 1);
            }

            Int32 rowCount = RowCount();
            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && predicate(row))
                    return row;
            }
            return null;
        }

        IXLRangeRow IXLRange.LastRowUsed(Func<IXLRangeRow, Boolean> predicate)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLRangeRow LastRowUsed(Func<IXLRangeRow, Boolean> predicate = null)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow IXLRange.LastRowUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate)
        {
            return LastRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRow IXLRange.LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate)
        {
            return LastRowUsed(options, predicate);
        }

        internal XLRangeRow LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null)
        {
            if (predicate == null)
            {
                Int32 lastRowUsed = Worksheet.Internals.CellsCollection.LastRowUsed(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber,
                    options);

                return lastRowUsed == 0 ? null : Row(lastRowUsed - RangeAddress.FirstAddress.RowNumber + 1);
            }

            Int32 rowCount = RowCount();
            for (Int32 ro = rowCount; ro >= 1; ro--)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && predicate(row))
                    return row;
            }
            return null;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRows IXLRange.RowsUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate)
        {
            return RowsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeRows IXLRange.RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate)
        {
            return RowsUsed(options, predicate);
        }

        internal XLRangeRows RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null)
        {
            XLRangeRows rows = new XLRangeRows();
            Int32 rowCount = RowCount();
            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                    rows.Add(row);
            }
            return rows;
        }

        IXLRangeRows IXLRange.RowsUsed(Func<IXLRangeRow, Boolean> predicate)
        {
            return RowsUsed(predicate);
        }

        internal XLRangeRows RowsUsed(Func<IXLRangeRow, Boolean> predicate = null)
        {
            return RowsUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeColumns IXLRange.ColumnsUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate)
        {
            return ColumnsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents, predicate);
        }

        IXLRangeColumns IXLRange.ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate)
        {
            return ColumnsUsed(options, predicate);
        }

        internal virtual XLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null)
        {
            XLRangeColumns columns = new XLRangeColumns();
            Int32 columnCount = ColumnCount();
            for (Int32 co = 1; co <= columnCount; co++)
            {
                var column = Column(co);

                if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                    columns.Add(column);
            }
            return columns;
        }

        IXLRangeColumns IXLRange.ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate)
        {
            return ColumnsUsed(predicate);
        }

        internal virtual XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate = null)
        {
            return ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        public XLRangeRow Row(Int32 row)
        {
            if (row <= 0 || row > XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
                throw new ArgumentOutOfRangeException(nameof(row), String.Format("Row number must be between 1 and {0}", XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1));

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

        public virtual XLRangeColumn Column(Int32 columnNumber)
        {
            if (columnNumber <= 0 || columnNumber > XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1)
                throw new ArgumentOutOfRangeException(nameof(columnNumber), String.Format("Column number must be between 1 and {0}", XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1));

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

        public virtual XLRangeColumn Column(String columnLetter)
        {
            return Column(XLHelper.GetColumnNumberFromLetter(columnLetter));
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
                    var newCell = new XLCell(Worksheet, newKey, oldCell.StyleValue);
                    newCell.CopyFrom(oldCell, true);
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
                    InsertColumnsAfter(false, rowCount - columnCount, false);
                else if (columnCount > rowCount)
                    InsertRowsBelow(false, columnCount - rowCount, false);
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
            var other = obj as XLRange;
            if (other == null)
                return false;
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
            Int32 columnCount = ColumnCount();
            for (Int32 c = 1; c <= columnCount; c++)
            {
                var column = Column(c);
                if (predicate == null || predicate(column))
                    return column;
            }
            return null;
        }

        public IXLRangeRow FindRow(Func<IXLRangeRow, bool> predicate)
        {
            Int32 rowCount = RowCount();
            for (Int32 r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate(row))
                    return row;
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
                return String.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    RangeAddress.FirstAddress.RowNumber,
                    ':',
                    RangeAddress.LastAddress.RowNumber);
            }
            else if (IsEntireColumn())
            {
                return String.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    RangeAddress.FirstAddress.ColumnLetter,
                    ':',
                    RangeAddress.LastAddress.ColumnLetter);
            }
            else
                return base.ToString();
        }
    }
}
