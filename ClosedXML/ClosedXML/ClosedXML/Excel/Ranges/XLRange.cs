using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRange: XLRangeBase, IXLRange
    {
        //public new IXLWorksheet Worksheet { get { return base.Worksheet; } }
        public XLRangeParameters RangeParameters { get; private set; }
        public XLRange(XLRangeParameters xlRangeParameters): base(xlRangeParameters.RangeAddress)
        {
            this.RangeParameters = xlRangeParameters;
            
            if (!xlRangeParameters.IgnoreEvents)
            {
                (Worksheet as XLWorksheet).RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
                (Worksheet as XLWorksheet).RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
                xlRangeParameters.IgnoreEvents = true;
            }
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(this.RangeAddress, range, columnsShifted);
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(this.RangeAddress, range, rowsShifted);
        }

        #region IXLRange Members

        public IXLRangeColumn FirstColumn()
        {
            return this.Column(1);
        }
        public IXLRangeColumn LastColumn()
        {
            return this.Column(this.ColumnCount());
        }
        public IXLRangeColumn FirstColumnUsed()
        {
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 minColumnUsed = Int32.MaxValue;
            Int32 minColumnInCells = Int32.MaxValue;
            if ((Worksheet as XLWorksheet).Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                minColumnInCells = (Worksheet as XLWorksheet).Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Min();

            Int32 minCoInColumns = Int32.MaxValue;
            if ((Worksheet as XLWorksheet).Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                minCoInColumns = (Worksheet as XLWorksheet).Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Min();

            minColumnUsed = minColumnInCells < minCoInColumns ? minColumnInCells : minCoInColumns;

            if (minColumnUsed == Int32.MaxValue)
                return null;
            else
                return this.Column(minColumnUsed);
        }
        public IXLRangeColumn LastColumnUsed()
        {
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 maxColumnUsed = 0;
            Int32 maxColumnInCells = 0;
            if ((Worksheet as XLWorksheet).Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                maxColumnInCells = (Worksheet as XLWorksheet).Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Max();

            Int32 maxCoInColumns = 0;
            if ((Worksheet as XLWorksheet).Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                maxCoInColumns = (Worksheet as XLWorksheet).Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Max();

            maxColumnUsed = maxColumnInCells > maxCoInColumns ? maxColumnInCells : maxCoInColumns;

            if (maxColumnUsed == 0)
                return null;
            else
                return this.Column(maxColumnUsed);
        }

        public IXLRangeRow FirstRow()
        {
            return this.Row(1);
        }
        public IXLRangeRow LastRow()
        {
            return this.Row(this.RowCount());
        }
        public IXLRangeRow FirstRowUsed()
        {
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
            var rowCount = this.RowCount();
            Int32 minRowUsed = Int32.MaxValue;
            Int32 minRowInCells = Int32.MaxValue;
            if ((Worksheet as XLWorksheet).Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                minRowInCells = (Worksheet as XLWorksheet).Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Min();

            Int32 minRoInRows = Int32.MaxValue;
            if ((Worksheet as XLWorksheet).Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                minRoInRows = (Worksheet as XLWorksheet).Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Min();

            minRowUsed = minRowInCells < minRoInRows ? minRowInCells : minRoInRows;

            if (minRowUsed == Int32.MaxValue)
                return null;
            else
                return this.Row(minRowUsed);
        }
        public IXLRangeRow LastRowUsed()
        {
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
            var rowCount = this.RowCount();
            Int32 maxRowUsed = 0;
            Int32 maxRowInCells = 0;
            if ((Worksheet as XLWorksheet).Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                maxRowInCells = (Worksheet as XLWorksheet).Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Max();

            Int32 maxRoInRows = 0;
            if ((Worksheet as XLWorksheet).Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                maxRoInRows = (Worksheet as XLWorksheet).Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Max();

            maxRowUsed = maxRowInCells > maxRoInRows ? maxRowInCells : maxRoInRows;

            if (maxRowUsed == 0)
                return null;
            else
                return this.Row(maxRowUsed);
        }

        public IXLRangeRow Row(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.FirstAddress.ColumnNumber, false, false);
            IXLAddress lastCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.LastAddress.ColumnNumber, false, false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style));
                
        }
        public IXLRangeRow RowQuick(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.FirstAddress.ColumnNumber, false, false);
            IXLAddress lastCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.LastAddress.ColumnNumber, false, false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress),Worksheet.Style), true);

        }
        public IXLRangeColumn Column(Int32 column)
        {
            IXLAddress firstCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            IXLAddress lastCellAddress = new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress),Worksheet.Style));
        }
        public IXLRangeColumn ColumnQuick(Int32 column)
        {
            IXLAddress firstCellAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            IXLAddress lastCellAddress = new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), true);
        }
        public IXLRangeColumn Column(String column)
        {
            return this.Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        public IXLRangeColumns Columns()
        {
            var retVal = new XLRangeColumns();
            foreach (var c in Enumerable.Range(1, this.ColumnCount()))
            {
                retVal.Add(this.Column(c));
            }
            return retVal;
        }
        public IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRangeColumns();

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add(this.Column(co));
            }
            return retVal;
        }
        public IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return this.Columns(XLAddress.GetColumnNumberFromLetter(firstColumn), XLAddress.GetColumnNumberFromLetter(lastColumn));
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
                        tPair = tPair.Replace('-', ':');

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
                    foreach (var col in this.Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                    {
                        retVal.Add(col);
                    }
                else
                    foreach (var col in this.Columns(firstColumn, lastColumn))
                    {
                        retVal.Add(col);
                    }
            }
            return retVal;
        }

        public IXLRangeRows Rows()
        {
            var retVal = new XLRangeRows();
            foreach (var r in Enumerable.Range(1, this.RowCount()))
            {
                retVal.Add(this.Row(r));
            }
            return retVal;
        }
        public IXLRangeRows Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRangeRows();

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(this.Row(ro));
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
                        tPair = tPair.Replace('-', ':');

                    var rowRange = tPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        public void Transpose(XLTransposeOptions transposeOption)
        {
            var rowCount = this.RowCount();
            var columnCount = this.ColumnCount();
            var squareSide = rowCount > columnCount ? rowCount : columnCount;

            var firstCell = FirstCell();
            var lastCell = LastCell();

            MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
            TransposeMerged(squareSide);
            TransposeRange(squareSide);
            this.RangeAddress.LastAddress = new XLAddress(Worksheet, 
                firstCell.Address.RowNumber + columnCount - 1,
                firstCell.Address.ColumnNumber + rowCount - 1,
                RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
            if (rowCount > columnCount)
            {
                var rng = Worksheet.Range(
                    this.RangeAddress.LastAddress.RowNumber + 1,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                    this.RangeAddress.LastAddress.ColumnNumber);
                rng.Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
            else if (columnCount > rowCount)
            {
                var rng = Worksheet.Range(
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + 1,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
                rng.Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }

            foreach (var c in this.Range(1,1,columnCount, rowCount).Cells())
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
            XLRange rngToTranspose = (XLRange)Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                this.RangeAddress.FirstAddress.ColumnNumber,
                this.RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                this.RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            Int32 roInitial = rngToTranspose.RangeAddress.FirstAddress.RowNumber;
            Int32 coInitial = rngToTranspose.RangeAddress.FirstAddress.ColumnNumber;
            Int32 roCount = rngToTranspose.RowCount();
            Int32 coCount = rngToTranspose.ColumnCount();
            for (Int32 ro = 1; ro <= roCount; ro++)
            {
                for (Int32 co = 1; co <= coCount; co++)
                {
                    var oldCell = rngToTranspose.Cell(ro, co);
                    var newKey = rngToTranspose.Cell(co, ro).Address; // new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
                    var newCell = new XLCell(newKey, oldCell.Style, Worksheet as XLWorksheet);
                    newCell.CopyFrom(oldCell);
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(oldCell.Address);
                }
            }
            //foreach (var c in rngToTranspose.Cells())
            //{
            //    var newKey = new XLAddress(Worksheet, c.Address.ColumnNumber, c.Address.RowNumber);
            //    var newCell = new XLCell(newKey, c.Style, Worksheet);
            //    newCell.Value = c.Value;
            //    newCell.DataType = c.DataType;
            //    cellsToInsert.Add(newKey, newCell);
            //    cellsToDelete.Add(c.Address);
            //}
            cellsToDelete.ForEach(c => (Worksheet as XLWorksheet).Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => (Worksheet as XLWorksheet).Internals.CellsCollection.Add(c.Key, c.Value));
        }

        private void TransposeMerged(Int32 squareSide)
        {
            XLRange rngToTranspose = (XLRange)Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                this.RangeAddress.FirstAddress.ColumnNumber,
                this.RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                this.RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            List<IXLRange> mergeToDelete = new List<IXLRange>();
            List<IXLRange> mergeToInsert = new List<IXLRange>();
            foreach (var merge in (Worksheet as XLWorksheet).Internals.MergedRanges)
            {
                if (this.Contains(merge))
                {
                    merge.RangeAddress.LastAddress = rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address;
                }
            }
            mergeToDelete.ForEach(m => (Worksheet as XLWorksheet).Internals.MergedRanges.Remove(m));
            mergeToInsert.ForEach(m => (Worksheet as XLWorksheet).Internals.MergedRanges.Add(m));
        }

        private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
        {
            if (transposeOption == XLTransposeOptions.MoveCells)
            {
                if (rowCount > columnCount)
                {
                    this.InsertColumnsAfter(rowCount - columnCount, false);
                }
                else if (columnCount > rowCount)
                {
                    this.InsertRowsBelow(columnCount - rowCount, false);
                }
            }
            else
            {
                if (rowCount > columnCount)
                {
                    var toMove = rowCount - columnCount;
                    var rngToClear = Worksheet.Range(
                        this.RangeAddress.FirstAddress.RowNumber,
                        this.RangeAddress.LastAddress.ColumnNumber + 1,
                        this.RangeAddress.LastAddress.RowNumber,
                        this.RangeAddress.LastAddress.ColumnNumber + toMove);
                    rngToClear.Clear();
                }
                else if (columnCount > rowCount)
                {
                    var toMove = columnCount - rowCount;
                    var rngToClear = Worksheet.Range(
                        this.RangeAddress.LastAddress.RowNumber + 1,
                        this.RangeAddress.FirstAddress.ColumnNumber,
                        this.RangeAddress.LastAddress.RowNumber + toMove,
                        this.RangeAddress.LastAddress.ColumnNumber);
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
            var other = (XLRange)obj;
            return this.RangeAddress.Equals(other.RangeAddress)
                && this.Worksheet.Equals(other.Worksheet);
        }

        public override int GetHashCode()
        {
            return RangeAddress.GetHashCode()
                    ^ this.Worksheet.GetHashCode();
        }

        IXLSortElements sortRows;
        public IXLSortElements SortRows
        {
            get
            {
                if (sortRows == null) sortRows = new XLSortElements();
                return sortRows;
            }
        }

        IXLSortElements sortColumns;
        public IXLSortElements SortColumns
        {
            get
            {
                if (sortColumns == null) sortColumns = new XLSortElements();
                return sortColumns;
            }
        }

        public IXLRange Sort()
        {
            if (SortColumns.Count() == 0)
                return Sort(XLSortOrder.Ascending);
            else
            {
                SortRangeRows();
                return this;
            }
        }
        public IXLRange Sort(Boolean matchCase)
        {
            if (SortColumns.Count() == 0)
                return Sort(XLSortOrder.Ascending, false);
            else
            {
                SortRangeRows();
                return this;
            }
        }
        public IXLRange Sort(XLSortOrder sortOrder)
        {
            if (SortColumns.Count() == 0)
            {
                Int32 columnCount = this.ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                    SortColumns.Add(co, sortOrder);
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
                Int32 columnCount = this.ColumnCount();
                for (Int32 co = 1; co <= columnCount; co++)
                    SortColumns.Add(co, sortOrder, true, matchCase);
            }
            else
            {
                SortColumns.ForEach(sc => { sc.SortOrder = sortOrder; sc.MatchCase = matchCase; });
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
                    co = XLAddress.GetColumnNumberFromLetter(coString);

                if (order.ToUpper().Equals("ASC"))
                    SortColumns.Add(co, XLSortOrder.Ascending);
                else
                    SortColumns.Add(co, XLSortOrder.Descending);
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
                    co = XLAddress.GetColumnNumberFromLetter(coString);

                if (order.ToUpper().Equals("ASC"))
                    SortColumns.Add(co, XLSortOrder.Ascending, true, matchCase);
                else
                    SortColumns.Add(co, XLSortOrder.Descending, true, matchCase);
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
            else
            {
                if (SortRows.Count() == 0)
                    return Sort(sortOrientation, XLSortOrder.Ascending);
                else
                {
                    SortRangeColumns();
                    return this;
                }
            }
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(matchCase);
            }
            else
            {
                if (SortRows.Count() == 0)
                    return Sort(sortOrientation, XLSortOrder.Ascending, matchCase);
                else
                {
                    SortRangeColumns();
                    return this;
                }
            }
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(sortOrder);
            }
            else
            {
                if (SortRows.Count() == 0)
                {
                    Int32 rowCount = this.RowCount();
                    for (Int32 co = 1; co <= rowCount; co++)
                        SortRows.Add(co, sortOrder);
                }
                else
                {
                    SortRows.ForEach(sc => sc.SortOrder = sortOrder);
                }
                SortRangeColumns();
                return this;
            }
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(sortOrder, matchCase);
            }
            else
            {
                if (SortRows.Count() == 0)
                {
                    Int32 rowCount = this.RowCount();
                    for (Int32 co = 1; co <= rowCount; co++)
                        SortRows.Add(co, sortOrder, matchCase);
                }
                else
                {
                    SortRows.ForEach(sc => { sc.SortOrder = sortOrder; sc.MatchCase = matchCase; });
                }
                SortRangeColumns();
                return this;
            }
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(elementsToSortBy);
            }
            else
            {
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
                        SortRows.Add(ro, XLSortOrder.Ascending);
                    else
                        SortRows.Add(ro, XLSortOrder.Descending);
                }

                SortRangeColumns();
                return this;
            }
        }
        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy, Boolean matchCase)
        {
            if (sortOrientation == XLSortOrientation.TopToBottom)
            {
                return Sort(elementsToSortBy, matchCase);
            }
            else
            {
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
                        SortRows.Add(ro, XLSortOrder.Ascending,true, matchCase);
                    else
                        SortRows.Add(ro, XLSortOrder.Descending, true, matchCase);
                }

                SortRangeColumns();
                return this;
            }
        }

        #region Sort Rows
        private void SortRangeRows()
        {
            SortingRangeRows(1, this.RowCount());
        }
        private void SwapRows(Int32 row1, Int32 row2)
        {

            Int32 cellCount = ColumnCount();

            for (Int32 co = 1; co <= cellCount; co++)
            {

                var cell1 = (XLCell)RowQuick(row1).Cell(co);
                var cell1Address = cell1.Address;
                var cell2 = (XLCell)RowQuick(row2).Cell(co);

                cell1.Address = cell2.Address;
                cell2.Address = cell1Address;

                (Worksheet as XLWorksheet).Internals.CellsCollection[cell1.Address] = cell1;
                (Worksheet as XLWorksheet).Internals.CellsCollection[cell2.Address] = cell2;
            }

        }
        private int SortRangeRows(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) &&
                   ((RowQuick(pivot) as XLRangeRow).CompareTo((RowQuick(m) as XLRangeRow), SortColumns) >= 0))
            {
                m++;
            }

            while ((n > begPoint) &&
                   ((RowQuick(pivot) as XLRangeRow).CompareTo((RowQuick(n) as XLRangeRow), SortColumns) <= 0))
            {
                n--;
            }
            while (m < n)
            {
                SwapRows(m, n);

                while ((m < endPoint) &&
                       ((RowQuick(pivot) as XLRangeRow).CompareTo((RowQuick(m) as XLRangeRow), SortColumns) >= 0))
                {
                    m++;
                }

                while ((n > begPoint) &&
                       ((RowQuick(pivot) as XLRangeRow).CompareTo((RowQuick(n) as XLRangeRow), SortColumns) <= 0))
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
            else
            {
                int pivot = SortRangeRows(beg, end);
                if (pivot > beg)
                    SortingRangeRows(beg, pivot - 1);
                if (pivot < end)
                    SortingRangeRows(pivot + 1, end);
            }
        }
        #endregion

        #region Sort Columns
        private void SortRangeColumns()
        {
            SortingRangeColumns(1, this.ColumnCount());
        }
        private void SwapColumns(Int32 column1, Int32 column2)
        {

            Int32 cellCount = ColumnCount();

            for (Int32 co = 1; co <= cellCount; co++)
            {

                var cell1 = (XLCell)ColumnQuick(column1).Cell(co);
                var cell1Address = cell1.Address;
                var cell2 = (XLCell)ColumnQuick(column2).Cell(co);

                cell1.Address = cell2.Address;
                cell2.Address = cell1Address;

                (Worksheet as XLWorksheet).Internals.CellsCollection[cell1.Address] = cell1;
                (Worksheet as XLWorksheet).Internals.CellsCollection[cell2.Address] = cell2;
            }

        }
        private int SortRangeColumns(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) &&
                   ((ColumnQuick(pivot) as XLRangeColumn).CompareTo((ColumnQuick(m) as XLRangeColumn), SortRows) >= 0))
            {
                m++;
            }

            while ((n > begPoint) &&
                   ((ColumnQuick(pivot) as XLRangeColumn).CompareTo((ColumnQuick(n) as XLRangeColumn), SortRows) <= 0))
            {
                n--;
            }
            while (m < n)
            {
                SwapColumns(m, n);

                while ((m < endPoint) &&
                       ((ColumnQuick(pivot) as XLRangeColumn).CompareTo((ColumnQuick(m) as XLRangeColumn), SortRows) >= 0))
                {
                    m++;
                }

                while ((n > begPoint) &&
                       ((ColumnQuick(pivot) as XLRangeColumn).CompareTo((ColumnQuick(n) as XLRangeColumn), SortRows) <= 0))
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
            else
            {
                int pivot = SortRangeColumns(beg, end);
                if (pivot > beg)
                    SortingRangeColumns(beg, pivot - 1);
                if (pivot < end)
                    SortingRangeColumns(pivot + 1, end);
            }
        }
        #endregion

        public new IXLRange CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            Int32 lastRowNumber = target.Address.RowNumber + this.RowCount() - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows) lastRowNumber = XLWorksheet.MaxNumberOfRows;
            Int32 lastColumnNumber = target.Address.ColumnNumber + this.ColumnCount() - 1;
            if (lastColumnNumber > XLWorksheet.MaxNumberOfColumns) lastColumnNumber = XLWorksheet.MaxNumberOfColumns;

            return target.Worksheet.Range(target.Address.RowNumber, target.Address.ColumnNumber,
                lastRowNumber, lastColumnNumber);
        }
        public new IXLRange CopyTo(IXLRangeBase target)
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
                lastColumnNumber);
        }

        public IXLRange SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }
    }
}
