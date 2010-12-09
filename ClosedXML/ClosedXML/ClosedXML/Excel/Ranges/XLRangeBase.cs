using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase : IXLRangeBase
    {
        public XLRangeBase(IXLRangeAddress rangeAddress)
        {
            RangeAddress = rangeAddress;
        }

        protected IXLStyle defaultStyle;
        public IXLRangeAddress RangeAddress { get; protected set; }
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

            if (cellsUsed.Count() == 0)
            {
                return null;
            }
            else
            {
                var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
                var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);
                return Worksheet.Cell(firstRow, firstColumn);
            }
        }

        public IXLCell LastCellUsed(Boolean ignoreStyle = true) 
        {
            var cellsUsed = CellsUsed();
            if (ignoreStyle)
                cellsUsed = cellsUsed.Where(c => c.GetString().Length != 0);

            if (cellsUsed.Count() == 0)
            {
                return null;
            }
            else
            {
                var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
                var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);
                return Worksheet.Cell(lastRow, lastColumn);
            }
        }
        
        public IXLCell Cell(Int32 row, Int32 column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell(String cellAddressInRange)
        {
            return this.Cell(new XLAddress(cellAddressInRange));
        }
        public IXLCell CellFast(String cellAddressInRange)
        {
            String toUse = cellAddressInRange;
            //if (cellAddressInRange.Contains('$'))
            //    toUse = cellAddressInRange.Replace("$", "");
            //else
            //    toUse = cellAddressInRange;

            var cIndex = 1;
            while (!Char.IsNumber(toUse, cIndex))
                cIndex++;

            return this.Cell(new XLAddress(Int32.Parse(toUse.Substring(cIndex)), toUse.Substring(0, cIndex)));
        }
        public IXLCell Cell(Int32 row, String column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell(IXLAddress cellAddressInRange)
        {
            return Cell((XLAddress)cellAddressInRange);
        }
        public IXLCell Cell(XLAddress cellAddressInRange)
        {
            IXLAddress absoluteAddress = cellAddressInRange + (XLAddress)this.RangeAddress.FirstAddress - 1;
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
            return this.RangeAddress.LastAddress.RowNumber - this.RangeAddress.FirstAddress.RowNumber + 1;
        }
        public Int32 ColumnCount()
        {
            return this.RangeAddress.LastAddress.ColumnNumber - this.RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLRange Range(String rangeAddressStr)
        {
            var rangeAddress = new XLRangeAddress(rangeAddressStr);
            return Range(rangeAddress);
        }
        public IXLRange Range(String firstCellAddress, String lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(firstCellAddress, lastCellAddress);
            return Range(rangeAddress);
        }
        public IXLRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            var rangeAddress = new XLRangeAddress(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
            return Range(rangeAddress);
        }
        public IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(firstCellAddress, lastCellAddress);
            return Range(rangeAddress);
        }
        public IXLRange Range(IXLRangeAddress rangeAddress)
        {
            var newFirstCellAddress = (XLAddress)rangeAddress.FirstAddress + (XLAddress)this.RangeAddress.FirstAddress - 1;
            var newLastCellAddress = (XLAddress)rangeAddress.LastAddress + (XLAddress)this.RangeAddress.FirstAddress - 1;
            var newRangeAddress = new XLRangeAddress(newFirstCellAddress, newLastCellAddress);
            var xlRangeParameters = new XLRangeParameters(newRangeAddress, this.Worksheet, this.Style);
            if (
                   newFirstCellAddress.RowNumber < this.RangeAddress.FirstAddress.RowNumber
                || newFirstCellAddress.RowNumber > this.RangeAddress.LastAddress.RowNumber
                || newLastCellAddress.RowNumber > this.RangeAddress.LastAddress.RowNumber
                || newFirstCellAddress.ColumnNumber < this.RangeAddress.FirstAddress.ColumnNumber
                || newFirstCellAddress.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber
                || newLastCellAddress.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber
                )
                throw new ArgumentOutOfRangeException(String.Format("The cells {0} and {1} are outside the range '{2}'.", newFirstCellAddress.ToString(), newLastCellAddress.ToString(), this.ToString()));

            return new XLRange(xlRangeParameters);
        }

        public IXLRanges Ranges( String ranges)
        {
            var retVal = new XLRanges(Worksheet.Internals.Workbook, Worksheet.Style);
            var rangePairs = ranges.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }
        public IXLRanges Ranges( params String[] ranges)
        {
            var retVal = new XLRanges(Worksheet.Internals.Workbook, Worksheet.Style);
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
            var list = this.Worksheet.Internals.CellsCollection.Where(c => this.ContainsRange(c.Key.ToString())).Select(c => c.Value);
            var retList = new List<IXLCell>();
            list.ForEach(c => retList.Add(c));
            return retList.AsEnumerable();
        }

        public IXLRange Merge()
        {
            var mergeRange = this.RangeAddress.FirstAddress.ToString() + ":" + this.RangeAddress.LastAddress.ToString();
            if (!this.Worksheet.Internals.MergedCells.Contains(mergeRange))
                this.Worksheet.Internals.MergedCells.Add(mergeRange);
            return AsRange();
        }
        public IXLRange Unmerge()
        {
            var mergeRange = this.RangeAddress.FirstAddress.ToString() + ":" + this.RangeAddress.LastAddress.ToString();
            if (this.Worksheet.Internals.MergedCells.Contains(mergeRange))
            this.Worksheet.Internals.MergedCells.Remove(mergeRange);
            return AsRange();
        }


        public void InsertColumnsAfter(Int32 numberOfColumns)
        {
            this.InsertColumnsAfter(numberOfColumns, false);
        }
        public void InsertColumnsAfter(Int32 numberOfColumns, Boolean onlyUsedCells)
        {
            var columnCount = this.ColumnCount();
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber + columnCount;
            if (firstColumn > XLWorksheet.MaxNumberOfColumns) firstColumn = XLWorksheet.MaxNumberOfColumns;
            var lastColumn = firstColumn + this.ColumnCount() - 1;
            if (lastColumn > XLWorksheet.MaxNumberOfColumns) lastColumn = XLWorksheet.MaxNumberOfColumns;

            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
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
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
            var lastRow = this.RangeAddress.FirstAddress.RowNumber + this.RowCount() - 1;

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
            var firstRow = this.RangeAddress.FirstAddress.RowNumber + rowCount;
            if (firstRow > XLWorksheet.MaxNumberOfRows) firstRow = XLWorksheet.MaxNumberOfRows;
            var lastRow = firstRow + this.RowCount() - 1;
            if (lastRow > XLWorksheet.MaxNumberOfRows) lastRow = XLWorksheet.MaxNumberOfRows;

            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
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
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = this.RangeAddress.FirstAddress.ColumnNumber + this.ColumnCount() - 1;

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
                    c.Address.ColumnNumber >= this.RangeAddress.FirstAddress.ColumnNumber
                    && c.Address.ColumnNumber <= this.RangeAddress.LastAddress.ColumnNumber
                    && c.Address.RowNumber >= this.RangeAddress.FirstAddress.RowNumber
                    && c.Address.RowNumber <= this.RangeAddress.LastAddress.RowNumber
                    );

            ClearMerged();
        }

        private void ClearMerged()
        {
            List<String> mergeToDelete = new List<String>();
            foreach (var merge in Worksheet.Internals.MergedCells)
            {
                var ma = new XLRangeAddress(merge);
                var ra = RangeAddress;

                if (!( // See if the two ranges intersect...
                       ma.FirstAddress.ColumnNumber > ra.LastAddress.ColumnNumber
                    || ma.LastAddress.ColumnNumber < ra.FirstAddress.ColumnNumber
                    || ma.FirstAddress.RowNumber > ra.LastAddress.RowNumber
                    || ma.LastAddress.RowNumber < ra.FirstAddress.RowNumber
                    ))
                {
                    mergeToDelete.Add(merge);
                }
            }
            mergeToDelete.ForEach(m => this.Worksheet.Internals.MergedCells.Remove(m));
        }

        public Boolean ContainsRange(String rangeAddress)
        {
            String addressToUse;
            if (rangeAddress.Contains("!"))
                addressToUse = rangeAddress.Substring(rangeAddress.IndexOf("!") + 1);
            else
                addressToUse = rangeAddress;

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                String[] arrRange = addressToUse.Split(':');
                firstAddress = new XLAddress(arrRange[0]);
                lastAddress = new XLAddress(arrRange[1]);
            }
            else
            {
                firstAddress = new XLAddress(addressToUse);
                lastAddress = new XLAddress(addressToUse);
            }
            return
                firstAddress >= (XLAddress)this.RangeAddress.FirstAddress
                && lastAddress <= (XLAddress)this.RangeAddress.LastAddress;
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            //this.Clear();

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var shiftLeftQuery = this.Worksheet.Internals.CellsCollection
                .Where(c =>
                       c.Key.RowNumber >= this.RangeAddress.FirstAddress.RowNumber
                    && c.Key.RowNumber <= this.RangeAddress.LastAddress.RowNumber
                    && c.Key.ColumnNumber >= this.RangeAddress.FirstAddress.ColumnNumber);

            var shiftUpQuery = this.Worksheet.Internals.CellsCollection
                .Where(c =>
                       c.Key.ColumnNumber >= this.RangeAddress.FirstAddress.ColumnNumber
                    && c.Key.ColumnNumber <= this.RangeAddress.LastAddress.ColumnNumber
                    && c.Key.RowNumber >= this.RangeAddress.FirstAddress.RowNumber);

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
                    c.Key.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber :
                    c.Key.RowNumber > this.RangeAddress.LastAddress.RowNumber;

                if (canInsert)
                    cellsToInsert.Add(newKey, newCell);
            }
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
            var shiftedRange = (XLRange)this.AsRange();
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
            }
            else
            {
                Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);
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
            return Worksheet.Range(RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("'");
            sb.Append(Worksheet.Name);
            sb.Append("'!");
            var firstAddress = new XLAddress(RangeAddress.FirstAddress.ToString());
            firstAddress.FixedColumn = true;
            firstAddress.FixedRow = true;
            sb.Append(firstAddress.ToString());
            sb.Append(":");
            var lastAddress = new XLAddress(RangeAddress.LastAddress.ToString());
            lastAddress.FixedColumn = true;
            lastAddress.FixedRow = true;
            sb.Append(lastAddress.ToString());
            return sb.ToString();
        }

        public String FormulaA1
        {
            set
            {
                Cells().ForEach(c => c.FormulaA1 = value);
            }
        }
        public String FormulaR1C1
        {
            set
            {
                Cells().ForEach(c => c.FormulaR1C1 = value);
            }
        }

        public IXLRange AddToNamed(String rangeName, XLScope scope = XLScope.Workbook, String comment = null)
        {
            IXLNamedRanges namedRanges;
            if (scope == XLScope.Workbook)
            {
                namedRanges = Worksheet.Internals.Workbook.NamedRanges;
            }
            else
            {
                namedRanges = Worksheet.NamedRanges;
            }

            if (namedRanges.Where(nr => nr.Name.ToLower() == rangeName.ToLower()).Any())
            {
                var namedRange = namedRanges.Where(nr => nr.Name.ToLower() == rangeName.ToLower()).Single();
                namedRange.Add(this.AsRange());
            }
            else
            {
                namedRanges.Add(rangeName, this.AsRange(), comment);
            }
            return AsRange();
        }

        protected void ShiftColumns(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int columnsShifted)
        {
            if (!thisRangeAddress.IsInvalid && !shiftedRange.RangeAddress.IsInvalid)
            {
                if ((columnsShifted < 0
                    // all columns
                    && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                    && thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.FirstAddress.ColumnNumber - columnsShifted
                    // all rows
                    && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
                    && thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber
                    ) || (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                        && shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber
                        && thisRangeAddress.FirstAddress.ColumnNumber + columnsShifted <= 0))
                {
                    thisRangeAddress.IsInvalid = true;
                }
                else
                {
                    if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                        && shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber)
                    {
                        if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber)
                            thisRangeAddress.FirstAddress = new XLAddress(thisRangeAddress.FirstAddress.RowNumber, thisRangeAddress.FirstAddress.ColumnNumber + columnsShifted);

                        if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.LastAddress.ColumnNumber)
                            thisRangeAddress.LastAddress = new XLAddress(thisRangeAddress.LastAddress.RowNumber, thisRangeAddress.LastAddress.ColumnNumber + columnsShifted);
                    }
                }
            }
        }

        protected void ShiftRows(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int rowsShifted)
        {
            if (!thisRangeAddress.IsInvalid && !shiftedRange.RangeAddress.IsInvalid)
            {
                if ((rowsShifted < 0
                    // all columns
                    && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                    && thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                    // all rows
                    && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
                    && thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber - rowsShifted
                    ) || (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                        && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber
                        && thisRangeAddress.FirstAddress.RowNumber + rowsShifted <= 0))
                {
                    thisRangeAddress.IsInvalid = true;
                }
                else
                {
                    if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                        && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber)
                    {
                        if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber)
                            thisRangeAddress.FirstAddress = new XLAddress(thisRangeAddress.FirstAddress.RowNumber + rowsShifted, thisRangeAddress.FirstAddress.ColumnNumber);

                        if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.LastAddress.RowNumber)
                            thisRangeAddress.LastAddress = new XLAddress(thisRangeAddress.LastAddress.RowNumber + rowsShifted, thisRangeAddress.LastAddress.ColumnNumber);
                    }
                }
            }
        }
    }
}
