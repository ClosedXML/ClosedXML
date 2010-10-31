using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase: IXLRangeBase
    {
        protected IXLStyle defaultStyle;
        public IXLAddress FirstAddressInSheet { get; protected set; }
        public IXLAddress LastAddressInSheet { get; protected set; }
        internal XLWorksheet Worksheet { get; set; }

        public IXLCell FirstCell()
        {
            return this.Cell(1, 1);
        }
        public IXLCell LastCell()
        {
            return this.Cell(this.RowCount(), this.ColumnCount());
        }

        public IXLCell FirstCellUsed(Boolean ignoreStyle = true)
        {
            var cellsUsed = CellsUsed();
            if (ignoreStyle)
                cellsUsed = cellsUsed.Where(c => c.GetString().Length != 0);

            var cellsUsedFiltered = cellsUsed.Where(cell => cell.Address == cellsUsed.Min(c => c.Address));
            
            if (cellsUsedFiltered.Count() > 0)
                return cellsUsedFiltered.Single();
            else
                return null;
        }

        public IXLCell LastCellUsed(Boolean ignoreStyle = true) 
        {
            var cellsUsed = CellsUsed();
            if (ignoreStyle)
                cellsUsed = cellsUsed.Where(c => c.GetString().Length != 0);

            var cellsUsedFiltered = cellsUsed.Where(cell => cell.Address == cellsUsed.Max(c => c.Address));
            if (cellsUsedFiltered.Count() > 0)
                return cellsUsedFiltered.Single();
            else
                return null;
        }
        
        public IXLCell Cell(Int32 row, Int32 column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell(String cellAddressInRange)
        {
            return this.Cell(new XLAddress(cellAddressInRange));
        }
        public IXLCell Cell(Int32 row, String column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell(IXLAddress cellAddressInRange)
        {
            IXLAddress absoluteAddress = (XLAddress)cellAddressInRange + (XLAddress)this.FirstAddressInSheet - 1;
            if (this.Worksheet.Internals.CellsCollection.ContainsKey(absoluteAddress))
            {
                return this.Worksheet.Internals.CellsCollection[absoluteAddress];
            }
            else
            {
                IXLStyle style = this.Style;
                if (this.Style.ToString() == this.Worksheet.Style.ToString())
                {
                    if (this.Worksheet.Internals.RowsCollection.ContainsKey(absoluteAddress.RowNumber)
                        && this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style.ToString() != this.Worksheet.Style.ToString())
                        style = this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style;
                    else if (this.Worksheet.Internals.ColumnsCollection.ContainsKey(absoluteAddress.ColumnNumber)
                        && this.Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style.ToString() != this.Worksheet.Style.ToString())
                        style = this.Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style;
                }
                var newCell = new XLCell(absoluteAddress, style, Worksheet);
                this.Worksheet.Internals.CellsCollection.Add(absoluteAddress, newCell);
                return newCell;
            }
        }

        public Int32 RowCount()
        {
            return this.LastAddressInSheet.RowNumber - this.FirstAddressInSheet.RowNumber + 1;
        }
        public Int32 ColumnCount()
        {
            return this.LastAddressInSheet.ColumnNumber - this.FirstAddressInSheet.ColumnNumber + 1;
        }

        public IXLRange Range( String rangeAddress)
        {
            if (rangeAddress.Contains(':'))
            {
                String[] arrRange = rangeAddress.Split(':');
                return this.Range(arrRange[0], arrRange[1]);
            }
            else
            {
                return this.Range(rangeAddress, rangeAddress);
            }
        }
        public IXLRange Range( String firstCellAddress, String lastCellAddress)
        {
            return this.Range(new XLAddress(firstCellAddress), new XLAddress(lastCellAddress));
        }
        public IXLRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            return this.Range(new XLAddress(firstCellRow, firstCellColumn), new XLAddress(lastCellRow, lastCellColumn));
        }
        public IXLRange Range( IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var newFirstCellAddress = (XLAddress)firstCellAddress + (XLAddress)this.FirstAddressInSheet - 1;
            var newLastCellAddress = (XLAddress)lastCellAddress + (XLAddress)this.FirstAddressInSheet - 1;
            var xlRangeParameters = new XLRangeParameters(newFirstCellAddress, newLastCellAddress, this.Worksheet, this.Style);
            if (
                   newFirstCellAddress.RowNumber < this.FirstAddressInSheet.RowNumber
                || newFirstCellAddress.RowNumber > this.LastAddressInSheet.RowNumber
                || newLastCellAddress.RowNumber > this.LastAddressInSheet.RowNumber
                || newFirstCellAddress.ColumnNumber < this.FirstAddressInSheet.ColumnNumber
                || newFirstCellAddress.ColumnNumber > this.LastAddressInSheet.ColumnNumber
                || newLastCellAddress.ColumnNumber > this.LastAddressInSheet.ColumnNumber
                )
                throw new ArgumentOutOfRangeException(String.Format("The cells {0} and {1} are outside the range '{2}'.", firstCellAddress.ToString(), lastCellAddress.ToString(), this.ToString()));

            return new XLRange(xlRangeParameters);
        }

        public IXLRanges Ranges( String ranges)
        {
            var retVal = new XLRanges(Worksheet);
            var rangePairs = ranges.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }
        public IXLRanges Ranges( params String[] ranges)
        {
            var retVal = new XLRanges(Worksheet);
            foreach (var pair in ranges)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }

        public IEnumerable<IXLCell> Cells()
        {
            foreach (var row in Enumerable.Range(1, this.RowCount()))
            {
                foreach (var column in Enumerable.Range(1, this.ColumnCount()))
                {
                    yield return this.Cell(row, column);
                }
            }
        }
        public IEnumerable<IXLCell> CellsUsed()
        {
            return this.Worksheet.Internals.CellsCollection.Values.AsEnumerable<IXLCell>();
        }

        public void Merge()
        {
            var mergeRange = this.FirstAddressInSheet.ToString() + ":" + this.LastAddressInSheet.ToString();
            if (!this.Worksheet.Internals.MergedCells.Contains(mergeRange))
                this.Worksheet.Internals.MergedCells.Add(mergeRange);
        }
        public void Unmerge()
        {
            this.Worksheet.Internals.MergedCells.Remove(this.FirstAddressInSheet.ToString() + ":" + this.LastAddressInSheet.ToString());
        }


        public void InsertColumnsAfter(Int32 numberOfColumns)
        {
            this.InsertColumnsAfter(numberOfColumns, false);
        }
        public void InsertColumnsAfter(Int32 numberOfColumns, Boolean onlyUsedCells)
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
        public void InsertColumnsBefore(Int32 numberOfColumns, Boolean onlyUsedCells)
        {
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
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
                        var newCell = new XLCell(newKey, oldCell.Style, Worksheet);
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
                    var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
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

            Worksheet.NotifyRangeShiftedColumns((XLRange)this.AsRange(), numberOfColumns);
        }

        public void InsertRowsBelow(Int32 numberOfRows)
        {
            this.InsertRowsBelow(numberOfRows, false);
        }
        public void InsertRowsBelow(Int32 numberOfRows, Boolean onlyUsedCells)
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
        public void InsertRowsAbove(Int32 numberOfRows, Boolean onlyUsedCells)
        {
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
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
                        var newCell = new XLCell(newKey, oldCell.Style, Worksheet);
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
                    var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
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

            Worksheet.NotifyRangeShiftedRows((XLRange)this.AsRange(), numberOfRows);
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
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
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
                var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
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
                Worksheet.NotifyRangeShiftedRows((XLRange)this.AsRange(), rowModifier * -1);
            }
            else
            {
                Worksheet.NotifyRangeShiftedColumns((XLRange)this.AsRange(), columnModifier * -1);
            }
        }

        #region IXLStylized Members

        public virtual IXLStyle Style
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

        public virtual IEnumerable<IXLStyle> Styles
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

        public virtual Boolean UpdatingStyle { get; set; }

        #endregion

        public virtual IXLRange AsRange()
        {
            return Worksheet.Range(FirstAddressInSheet, LastAddressInSheet);
        }

        public override string ToString()
        {
            return FirstAddressInSheet.ToString() + ":" + LastAddressInSheet.ToString();
        }

    }
}
