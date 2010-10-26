using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRange : XLRangeBase, IXLRange
    {
        public IXLStyle defaultStyle;

        public XLRange(XLRangeParameters xlRangeParameters)
        {
            FirstAddressInSheet = xlRangeParameters.FirstCellAddress;
            LastAddressInSheet = xlRangeParameters.LastCellAddress;
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            //Worksheet.Internals.RowsCollection.RowShifted += new RowShiftedDelegate(RowsCollection_RowShifted);
            //Worksheet.Internals.ColumnsCollection.ColumnShifted += new ColumnShiftedDelegate(ColumnsCollection_ColumnShifted);
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.FirstAddressInSheet.RowNumber <= FirstAddressInSheet.RowNumber
                && range.LastAddressInSheet.RowNumber >= LastAddressInSheet.RowNumber)
            {
                ColumnsCollection_ColumnShifted(range.FirstAddressInSheet.ColumnNumber, columnsShifted);
            }
        }

        void RowsCollection_RowShifted(int startingRow, int rowsShifted)
        {
            if (startingRow <= FirstAddressInSheet.RowNumber)
            {
                FirstAddressInSheet = new XLAddress(FirstAddressInSheet.RowNumber + rowsShifted, FirstAddressInSheet.ColumnNumber);
            }
            
            if (startingRow <= LastAddressInSheet.RowNumber)
            {
                LastAddressInSheet = new XLAddress(LastAddressInSheet.RowNumber + rowsShifted, LastAddressInSheet.ColumnNumber);
            }
        }

        void ColumnsCollection_ColumnShifted(int startingColumn, int columnsShifted)
        {
            if (startingColumn <= FirstAddressInSheet.ColumnNumber)
            {
                FirstAddressInSheet = new XLAddress(FirstAddressInSheet.RowNumber, FirstAddressInSheet.ColumnNumber + columnsShifted);
            }

            if (startingColumn <= LastAddressInSheet.ColumnNumber)
            {
                LastAddressInSheet = new XLAddress(LastAddressInSheet.RowNumber, LastAddressInSheet.ColumnNumber + columnsShifted);
            }
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.FirstAddressInSheet.ColumnNumber <= FirstAddressInSheet.ColumnNumber
                && range.LastAddressInSheet.ColumnNumber >= LastAddressInSheet.ColumnNumber)
            {
                RowsCollection_RowShifted(range.FirstAddressInSheet.RowNumber, rowsShifted);
            }
        }


        #region IXLRange Members

        public IXLRange FirstColumn()
        {
            return this.Column(1);
        }
        public IXLRange LastColumn()
        {
            return this.Column(this.ColumnCount());
        }
        public IXLRange FirstColumnUsed()
        {
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 minColumnUsed = Int32.MaxValue;
            Int32 minColumnInCells = Int32.MaxValue;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                minColumnInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Min();

            Int32 minCoInColumns = Int32.MaxValue;
            if (this.Worksheet.Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                minCoInColumns = this.Worksheet.Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Min();

            minColumnUsed = minColumnInCells < minCoInColumns ? minColumnInCells : minCoInColumns;

            if (minColumnUsed == Int32.MaxValue)
                return null;
            else
                return this.Row(minColumnUsed);
        }
        public IXLRange LastColumnUsed()
        {
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 maxColumnUsed = 0;
            Int32 maxColumnInCells = 0;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                maxColumnInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Max();

            Int32 maxCoInColumns = 0;
            if (this.Worksheet.Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                maxCoInColumns = this.Worksheet.Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Max();

            maxColumnUsed = maxColumnInCells > maxCoInColumns ? maxColumnInCells : maxCoInColumns;

            if (maxColumnUsed == 0)
                return null;
            else
                return this.Column(maxColumnUsed);
        }

        public IXLRange FirstRow()
        {
            return this.Row(1);
        }
        public IXLRange LastRow()
        {
            return this.Row(this.RowCount());
        }
        public IXLRange FirstRowUsed()
        {
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var rowCount = this.RowCount();
            Int32 minRowUsed = Int32.MaxValue;
            Int32 minRowInCells = Int32.MaxValue;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                minRowInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Min();

            Int32 minRoInRows = Int32.MaxValue;
            if (this.Worksheet.Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                minRoInRows = this.Worksheet.Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Min();

            minRowUsed = minRowInCells < minRoInRows ? minRowInCells : minRoInRows;

            if (minRowUsed == Int32.MaxValue)
                return null;
            else
                return this.Row(minRowUsed);
        }
        public IXLRange LastRowUsed()
        {
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var rowCount = this.RowCount();
            Int32 maxRowUsed = 0;
            Int32 maxRowInCells = 0;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                maxRowInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Max();

            Int32 maxRoInRows = 0;
            if (this.Worksheet.Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                maxRoInRows = this.Worksheet.Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Max();

            maxRowUsed = maxRowInCells > maxRoInRows ? maxRowInCells : maxRoInRows;

            if (maxRowUsed == 0)
                return null;
            else
                return this.Row(maxRowUsed);
        }

        public IXLRange Row(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(row, 1);
            IXLAddress lastCellAddress = new XLAddress(row, this.ColumnCount());
            return this.Range(firstCellAddress, lastCellAddress);
        }
        public IXLRange Column(Int32 column)
        {
            return this.Range(1, column, this.RowCount(), column);
        }
        public IXLRange Column(String column)
        {
            return this.Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        public IXLRanges Columns()
        {
            var retVal = new XLRanges();
            foreach (var c in Enumerable.Range(1, this.ColumnCount()))
            {
                retVal.Add(this.Column(c));
            }
            return retVal;
        }
        public IXLRanges Columns(String columns)
        {
            var retVal = new XLRanges();
            var columnPairs = columns.Split(',');
            foreach (var pair in columnPairs)
            {
                String firstColumn;
                String lastColumn;
                if (pair.Contains(':'))
                {
                    var columnRange = pair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = pair;
                    lastColumn = pair;
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
        public IXLRanges Columns(String firstColumn, String lastColumn)
        {
            return this.Columns(XLAddress.GetColumnNumberFromLetter(firstColumn), XLAddress.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLRanges Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRanges();

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add(this.Column(co));
            }
            return retVal;
        }
        public IXLRanges Rows()
        {
            var retVal = new XLRanges();
            foreach (var r in Enumerable.Range(1, this.RowCount()))
            {
                retVal.Add(this.Row(r));
            }
            return retVal;
        }
        public IXLRanges Rows(String rows)
        {
            var retVal = new XLRanges();
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
            {
                String firstRow;
                String lastRow;
                if (pair.Contains(':'))
                {
                    var rowRange = pair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = pair;
                    lastRow = pair;
                }
                foreach (var row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }
        public IXLRanges Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRanges();

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(this.Row(ro));
            }
            return retVal;
        }

        public void Clear()
        {
            // Remove cells inside range
            this.Worksheet.Internals.CellsCollection.RemoveAll(c =>
                    c.Address.ColumnNumber >= this.FirstAddressInSheet.ColumnNumber
                    && c.Address.ColumnNumber <= this.LastAddressInSheet.ColumnNumber
                    && c.Address.RowNumber >= this.FirstAddressInSheet.RowNumber
                    && c.Address.RowNumber <= this.LastAddressInSheet.RowNumber
                    );
        }
        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            //this.Clear();

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, IXLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var shiftLeftQuery = this.Worksheet.Internals.CellsCollection
                .Where(c =>  
                       c.Key.RowNumber >= this.FirstAddressInSheet.RowNumber
                    && c.Key.RowNumber <= this.LastAddressInSheet.RowNumber
                    && c.Key.ColumnNumber >= this.FirstAddressInSheet.ColumnNumber);

            var shiftUpQuery = this.Worksheet.Internals.CellsCollection
                .Where(c =>
                       c.Key.ColumnNumber >= this.FirstAddressInSheet.ColumnNumber
                    && c.Key.ColumnNumber <= this.LastAddressInSheet.ColumnNumber
                    && c.Key.RowNumber >= this.FirstAddressInSheet.RowNumber);

            var columnModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? this.ColumnCount() : 0;
            var rowModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp ? this.RowCount() : 0;
            var cellsQuery = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? shiftLeftQuery : shiftUpQuery;
            foreach (var c in cellsQuery)
            {
                var newKey = new XLAddress(c.Key.RowNumber - rowModifier, c.Key.ColumnNumber - columnModifier);
                var newCell = new XLCell(newKey, c.Value.Style);
                newCell.Value = c.Value.Value;
                newCell.DataType = c.Value.DataType;
                cellsToDelete.Add(c.Key);

                var canInsert = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ?
                    c.Key.ColumnNumber > this.LastAddressInSheet.ColumnNumber :
                    c.Key.RowNumber > this.LastAddressInSheet.RowNumber;

                if (canInsert)
                    cellsToInsert.Add(newKey, newCell);
            }
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                Worksheet.NotifyRangeShiftedRows(this, rowModifier * -1);
            }
            else
            {
                Worksheet.NotifyRangeShiftedColumns(this, columnModifier * -1);
            }
        }

        public void InsertRowsBelow(Int32 numberOfRows)
        {
            this.InsertRowsBelow(numberOfRows, false);
        }
        internal void InsertRowsBelow(Int32 numberOfRows, Boolean onlyUsedCells)
        {
            var rowCount = this.RowCount();
            var firstRow = this.FirstAddressInSheet.RowNumber + rowCount;
            if (firstRow > XLWorksheet.MaxNumberOfRows) firstRow = XLWorksheet.MaxNumberOfRows;
            var lastRow = firstRow + this.RowCount() - 1;
            if (lastRow > XLWorksheet.MaxNumberOfRows) lastRow = XLWorksheet.MaxNumberOfRows;

            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var lastColumn = firstColumn + this.ColumnCount() - 1;
            if (lastColumn > XLWorksheet.MaxNumberOfColumns) lastColumn = XLWorksheet.MaxNumberOfColumns;

            var newRange = (XLRange)this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            newRange.InsertRowsAbove(numberOfRows, onlyUsedCells);
        }
        public void InsertRowsAbove(Int32 numberOfRows)
        {
            this.InsertRowsAbove(numberOfRows, false);
        }
        internal void InsertRowsAbove(Int32 numberOfRows, Boolean onlyUsedCells)
        {
            var cellsToInsert = new Dictionary<IXLAddress, IXLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var cellsToBlank = new List<IXLAddress>();
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var lastColumn = this.FirstAddressInSheet.ColumnNumber + this.ColumnCount() - 1;

            if (!onlyUsedCells)
            {
                var lastRow = this.Worksheet.LastRowUsed().RowNumber();
                for (var ro = lastRow; ro >= firstRow; ro--)
                {
                    for (var co = lastColumn; co >= firstColumn; co--)
                    {
                        var oldKey = new XLAddress(ro, co);
                        var newRow = ro + numberOfRows;
                        var newKey = new XLAddress(newRow, co);
                        IXLCell oldCell;
                        if (this.Worksheet.Internals.CellsCollection.ContainsKey(oldKey))
                        {
                            oldCell = this.Worksheet.Internals.CellsCollection[oldKey];
                        }
                        else
                        {
                            oldCell = this.Worksheet.Cell(oldKey);
                        }
                        var newCell = new XLCell(newKey, oldCell.Style);
                        newCell.Value = oldCell.Value;
                        newCell.DataType = oldCell.DataType;
                        cellsToInsert.Add(newKey, newCell);
                        cellsToDelete.Add(oldKey);
                        if (oldKey.RowNumber < firstRow + numberOfRows)
                            cellsToBlank.Add(oldKey);
                    }
                }
            }
            else
            {
                foreach (var c in this.Worksheet.Internals.CellsCollection
                    .Where(c =>
                    c.Key.RowNumber >= firstRow
                    && c.Key.ColumnNumber >= firstColumn
                    && c.Key.ColumnNumber <= lastColumn
                    ))
                {
                    var newRow = c.Key.RowNumber + numberOfRows;
                    var newKey = new XLAddress(newRow, c.Key.ColumnNumber);
                    var newCell = new XLCell(newKey, c.Value.Style);
                    newCell.Value = c.Value.Value;
                    newCell.DataType = c.Value.DataType;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Key);
                    if (c.Key.RowNumber < firstRow + numberOfRows)
                        cellsToBlank.Add(c.Key);
                }
            }
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
            foreach (var c in cellsToBlank)
            {
                IXLStyle styleToUse;
                if (this.Worksheet.Internals.ColumnsCollection.ContainsKey(c.ColumnNumber))
                    styleToUse = this.Worksheet.Internals.ColumnsCollection[c.ColumnNumber].Style;
                else
                    styleToUse = this.Worksheet.Style;
                this.Worksheet.Cell(c).Style = styleToUse;
            }

            Worksheet.NotifyRangeShiftedRows(this, numberOfRows);
        }

        public void InsertColumnsAfter(Int32 numberOfColumns)
        {
            this.InsertColumnsAfter(numberOfColumns, false);
        }
        internal void InsertColumnsAfter(Int32 numberOfColumns, Boolean onlyUsedCells)
        {
            var columnCount = this.ColumnCount();
            var firstColumn = this.FirstAddressInSheet.ColumnNumber + columnCount;
            if (firstColumn > XLWorksheet.MaxNumberOfColumns) firstColumn = XLWorksheet.MaxNumberOfColumns;
            var lastColumn = firstColumn + this.ColumnCount() - 1;
            if (lastColumn > XLWorksheet.MaxNumberOfColumns) lastColumn = XLWorksheet.MaxNumberOfColumns;

            var firstRow = this.FirstAddressInSheet.RowNumber;
            var lastRow = firstRow + this.RowCount() - 1;
            if (lastRow > XLWorksheet.MaxNumberOfRows) lastRow = XLWorksheet.MaxNumberOfRows;

            var newRange = (XLRange)this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            newRange.InsertColumnsBefore(numberOfColumns, onlyUsedCells);
        }
        public void InsertColumnsBefore(Int32 numberOfColumns)
        {
            this.InsertColumnsBefore(numberOfColumns, false);
        }
        internal void InsertColumnsBefore(Int32 numberOfColumns, Boolean onlyUsedCells)
        {
            var cellsToInsert = new Dictionary<IXLAddress, IXLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var cellsToBlank = new List<IXLAddress>();
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var lastRow = this.FirstAddressInSheet.RowNumber + this.RowCount() - 1;

            if (!onlyUsedCells)
            {
                var lastColumn = this.Worksheet.LastColumnUsed().ColumnNumber();
                for (var co = lastColumn; co >= firstColumn; co--)
                {
                    for (var ro = lastRow; ro >= firstRow; ro--)
                    {
                        var oldKey = new XLAddress(ro, co);
                        var newColumn = co + numberOfColumns;
                        var newKey = new XLAddress(ro, newColumn);
                        IXLCell oldCell;
                        if (this.Worksheet.Internals.CellsCollection.ContainsKey(oldKey))
                        {
                            oldCell = this.Worksheet.Internals.CellsCollection[oldKey];
                        }
                        else
                        {
                            oldCell = this.Worksheet.Cell(oldKey);
                        }
                        var newCell = new XLCell(newKey, oldCell.Style);
                        newCell.Value = oldCell.Value;
                        newCell.DataType = oldCell.DataType;
                        cellsToInsert.Add(newKey, newCell);
                        cellsToDelete.Add(oldKey);
                        if (oldKey.ColumnNumber < firstColumn + numberOfColumns)
                            cellsToBlank.Add(oldKey);
                    }
                }
            }
            else
            {
                foreach (var c in this.Worksheet.Internals.CellsCollection
                    .Where(c =>
                    c.Key.ColumnNumber >= firstColumn
                    && c.Key.RowNumber >= firstRow
                    && c.Key.RowNumber <= lastRow
                    ))
                {
                    var newColumn = c.Key.ColumnNumber + numberOfColumns;
                    var newKey = new XLAddress(c.Key.RowNumber, newColumn);
                    var newCell = new XLCell(newKey, c.Value.Style);
                    newCell.Value = c.Value.Value;
                    newCell.DataType = c.Value.DataType;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Key);
                    if (c.Key.ColumnNumber < firstColumn + numberOfColumns)
                        cellsToBlank.Add(c.Key);
                }
            }
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
            foreach (var c in cellsToBlank)
            {
                IXLStyle styleToUse;
                if (this.Worksheet.Internals.RowsCollection.ContainsKey(c.RowNumber))
                    styleToUse = this.Worksheet.Internals.RowsCollection[c.RowNumber].Style;
                else
                    styleToUse = this.Worksheet.Style;
                this.Worksheet.Cell(c).Style = styleToUse;
            }

            Worksheet.NotifyRangeShiftedColumns(this, numberOfColumns);
        }
        
        #endregion

        #region IXLStylized Members

        public override IXLStyle Style 
        {
            get
            {
                return this.defaultStyle;
            }
            set
            {
                this.Cells().ForEach(c => c.Style = value);
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                foreach (var cell in this.Cells())
                {
                    yield return cell.Style;
                }
                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        #endregion

        public override IXLRange AsRange()
        {
            return this;
        }

        public override string ToString()
        {
            return FirstAddressInSheet.ToString() + ":" + LastAddressInSheet.ToString();
        }

    }
}
