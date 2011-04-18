using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase : IXLRangeBase, IXLStylized
    {
        public XLRangeBase(IXLRangeAddress rangeAddress)
        {
            RangeAddress = rangeAddress;
        }

        protected IXLStyle defaultStyle;
        public IXLRangeAddress RangeAddress { get; protected set; }
        public XLWorksheet Worksheet { get; set; }

        public IXLCell FirstCell()
        {
            return this.Cell(1, 1);
        }
        public IXLCell LastCell()
        {
            return this.Cell(this.RowCount(), this.ColumnCount());
        }
        
        public IXLCell FirstCellUsed()
        {
            return FirstCellUsed(false);
        }
        public IXLCell FirstCellUsed(Boolean includeStyles)
        {
            var cellsUsed = CellsUsed(includeStyles);

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

        public IXLCell LastCellUsed()
        {
            return LastCellUsed(false);
        }
        public IXLCell LastCellUsed(Boolean includeStyles) 
        {
            var cellsUsed = CellsUsed(includeStyles);

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
            return this.Cell(new XLAddress(row, column, false, false));
        }
        public IXLCell Cell(String cellAddressInRange)
        {
            return this.Cell(new XLAddress(cellAddressInRange));
        }

        public IXLCell Cell(Int32 row, String column)
        {
            return this.Cell(new XLAddress(row, column, false, false));
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
                if (this.Style != null && this.Style.Equals(this.Worksheet.Style))
                {
                    if (this.Worksheet.Internals.RowsCollection.ContainsKey(absoluteAddress.RowNumber)
                        && !this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style.Equals(this.Worksheet.Style))
                        style = this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style;
                    else if (this.Worksheet.Internals.ColumnsCollection.ContainsKey(absoluteAddress.ColumnNumber)
                        && !this.Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style.Equals(this.Worksheet.Style))
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
        public Int32 RowNumber()
        {
            return this.RangeAddress.FirstAddress.RowNumber;
        }
        public Int32 ColumnCount()
        {
            return this.RangeAddress.LastAddress.ColumnNumber - this.RangeAddress.FirstAddress.ColumnNumber + 1;
        }
        public Int32 ColumnNumber()
        {
            return this.RangeAddress.FirstAddress.ColumnNumber;
        }
        public String ColumnLetter()
        {
            return this.RangeAddress.FirstAddress.ColumnLetter;
        }

        public virtual IXLRange Range(String rangeAddressStr)
        {
            var rangeAddress = new XLRangeAddress(rangeAddressStr);
            return Range(rangeAddress);
        }

        public IXLRange Range(IXLCell firstCell, IXLCell lastCell)
        {
            return Range(firstCell.Address, lastCell.Address);
        }
        public IXLRange Range(String firstCellAddress, String lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(firstCellAddress, lastCellAddress);
            return Range(rangeAddress);
        }
        public IXLRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            var rangeAddress = new XLRangeAddress(new XLAddress(firstCellRow, firstCellColumn, false, false), new XLAddress(lastCellRow, lastCellColumn, false, false));
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
            newFirstCellAddress.FixedRow = rangeAddress.FirstAddress.FixedRow;
            newFirstCellAddress.FixedColumn = rangeAddress.FirstAddress.FixedColumn;
            var newLastCellAddress = (XLAddress)rangeAddress.LastAddress + (XLAddress)this.RangeAddress.FirstAddress - 1;
            newLastCellAddress.FixedRow = rangeAddress.LastAddress.FixedRow;
            newLastCellAddress.FixedColumn = rangeAddress.LastAddress.FixedColumn;
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
                retVal.Add(Range(pair.Trim()));
            }
            return retVal;
        }
        public IXLRanges Ranges(params String[] ranges)
        {
            var retVal = new XLRanges(Worksheet.Internals.Workbook, Worksheet.Style);
            foreach (var pair in ranges)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }
        protected String FixColumnAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
                return "A" + address;
            else
                return address;
        }
        protected String FixRowAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
                return XLAddress.GetColumnLetterFromNumber(test) + "1";
            else
                return address;
        }
        public IXLCells Cells()
        {
            var cells = new XLCells(Worksheet, false, false, false);
            cells.Add(this.RangeAddress);
            return (IXLCells)cells;
        }
        public IXLCells CellsUsed()
        {
            var cells = new XLCells(Worksheet, false, true, false);
            cells.Add(this.RangeAddress);
            return (IXLCells)cells;
        }
        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(Worksheet, false, true, includeStyles);
            cells.Add(this.RangeAddress);
            return (IXLCells)cells;
        }

        public IXLRange Merge()
        {
            var tAddress = this.RangeAddress.ToString();
            Boolean foundOne = false;
            foreach (var m in this.Worksheet.Internals.MergedRanges)
            {
                var mAddress = m.RangeAddress.ToString();
                if (mAddress == tAddress)
                {
                    foundOne = true;
                    break;
                }
            }

            if (!foundOne)
                this.Worksheet.Internals.MergedRanges.Add(this.AsRange());
            return AsRange();
        }
        public IXLRange Unmerge()
        {
            var tAddress = this.RangeAddress.ToString();
            foreach (var m in this.Worksheet.Internals.MergedRanges)
            {
                var mAddress = m.RangeAddress.ToString();
                if (mAddress == tAddress)
                {
                    this.Worksheet.Internals.MergedRanges.Remove(this.AsRange());
                    break;
                }
            }
                
            return AsRange();
        }

        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            return InsertColumnsAfter(numberOfColumns, true);
        }
        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = this.InsertColumnsAfter(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                this.RangeAddress = new XLRangeAddress(
                    new XLAddress(RangeAddress.FirstAddress.RowNumber,RangeAddress.FirstAddress.ColumnNumber, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber + numberOfColumns, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeColumns InsertColumnsAfter(Boolean onlyUsedCells, Int32 numberOfColumns)
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
            return newRange.InsertColumnsBefore(onlyUsedCells, numberOfColumns);
        }
        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            return InsertColumnsBefore(numberOfColumns, false);
        }
        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = this.InsertColumnsBefore(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                this.RangeAddress = new XLRangeAddress(
                new XLAddress(RangeAddress.FirstAddress.RowNumber,RangeAddress.FirstAddress.ColumnNumber - numberOfColumns, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeColumns InsertColumnsBefore(Boolean onlyUsedCells, Int32 numberOfColumns)
        {
            foreach (var ws in Worksheet.Internals.Workbook.Worksheets)
            {
                var xlWorksheet = (XLWorksheet)ws;
                foreach (var cell in xlWorksheet.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaColumns((XLRange)this.AsRange(), numberOfColumns);
                }
            }

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
                        var oldKey = new XLAddress(ro, co, false, false);
                        var newColumn = co + numberOfColumns;
                        var newKey = new XLAddress(ro, newColumn, false, false);
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
                        newCell.CopyValues((XLCell)oldCell);
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
                    var newKey = new XLAddress(c.Key.RowNumber, newColumn, false, false);
                    var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
                    newCell.CopyValues(c.Value);
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
            return Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber - numberOfColumns
                ).Columns();
        }

        public IXLRangeRows InsertRowsBelow(Int32 numberOfRows)
        {
            return InsertRowsBelow(numberOfRows, true);
        }
        public IXLRangeRows InsertRowsBelow(Int32 numberOfRows, Boolean expandRange)
        {
            var retVal = this.InsertRowsBelow(false,numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                this.RangeAddress = new XLRangeAddress(
                 new XLAddress(RangeAddress.FirstAddress.RowNumber,RangeAddress.FirstAddress.ColumnNumber, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(RangeAddress.LastAddress.RowNumber + numberOfRows, RangeAddress.LastAddress.ColumnNumber, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeRows InsertRowsBelow(Boolean onlyUsedCells, Int32 numberOfRows)
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
            return newRange.InsertRowsAbove(onlyUsedCells, numberOfRows);
        }
        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows)
        {
            return InsertRowsAbove(numberOfRows, false);
        }
        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows, Boolean expandRange)
        {
            var retVal = this.InsertRowsAbove(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                this.RangeAddress = new XLRangeAddress(
                new XLAddress(RangeAddress.FirstAddress.RowNumber - numberOfRows,RangeAddress.FirstAddress.ColumnNumber, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeRows InsertRowsAbove(Boolean onlyUsedCells, Int32 numberOfRows )
        {
            foreach (var ws in Worksheet.Internals.Workbook.Worksheets)
            {
                var xlWorksheet = (XLWorksheet)ws;
                foreach (var cell in xlWorksheet.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaRows((XLRange)this.AsRange(), numberOfRows);
                }
            }

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
                        var oldKey = new XLAddress(ro, co, false, false);
                        var newRow = ro + numberOfRows;
                        var newKey = new XLAddress(newRow, co, false, false);
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
                        newCell.CopyFrom(oldCell);
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
                    var newKey = new XLAddress(newRow, c.Key.ColumnNumber, false, false);
                    var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
                    newCell.CopyFrom(c.Value);
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
            return Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber - numberOfRows, 
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber - numberOfRows,
                RangeAddress.LastAddress.ColumnNumber
                ).Rows();
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

            List<XLHyperlink> hyperlinksToRemove = new List<XLHyperlink>();
            foreach (var hl in Worksheet.Hyperlinks)
            {
                if (this.Contains(hl.Cell.AsRange()))
                    hyperlinksToRemove.Add(hl);
            }
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));
        }

        public void ClearStyles()
        {
            foreach (var cell in CellsUsed(true))
            {
                var newStyle = new XLStyle((XLCell)cell, Worksheet.Style);
                newStyle.NumberFormat = cell.Style.NumberFormat;
                cell.Style = newStyle;
            }
        }

        private void ClearMerged()
        {
            List<IXLRange> mergeToDelete = new List<IXLRange>();
            foreach (var merge in Worksheet.Internals.MergedRanges)
            {
                if (this.Intersects(merge))
                {
                    mergeToDelete.Add(merge);
                }
            }
            mergeToDelete.ForEach(m => this.Worksheet.Internals.MergedRanges.Remove(m));
        }

        public Boolean Contains(String rangeAddress)
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

        public Boolean Contains(IXLRangeBase range)
        {
            return
                (XLAddress)range.RangeAddress.FirstAddress >= (XLAddress)this.RangeAddress.FirstAddress
                && (XLAddress)range.RangeAddress.LastAddress <= (XLAddress)this.RangeAddress.LastAddress;
        }

        public Boolean Intersects(String rangeAddress)
        {
            return this.Intersects(Range(rangeAddress));
        }

        public Boolean Intersects(IXLRangeBase range)
        {
            var ma = range.RangeAddress;
            var ra = RangeAddress;

            return !( // See if the two ranges intersect...
                   ma.FirstAddress.ColumnNumber > ra.LastAddress.ColumnNumber
                || ma.LastAddress.ColumnNumber < ra.FirstAddress.ColumnNumber
                || ma.FirstAddress.RowNumber > ra.LastAddress.RowNumber
                || ma.LastAddress.RowNumber < ra.FirstAddress.RowNumber
                );
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            var numberOfRows = this.RowCount();
            var numberOfColumns = this.ColumnCount();
            IXLRange shiftedRangeFormula;
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                var lastCell = Worksheet.Cell(XLWorksheet.MaxNumberOfRows, RangeAddress.LastAddress.ColumnNumber);
                shiftedRangeFormula = Worksheet.Range(RangeAddress.FirstAddress, lastCell.Address);
                if (StringExtensions.IsNullOrWhiteSpace(lastCell.GetString()) && StringExtensions.IsNullOrWhiteSpace(lastCell.FormulaA1))
                    Worksheet.Internals.CellsCollection.Remove(lastCell.Address);
            }
            else
            {
                var lastCell = Worksheet.Cell(RangeAddress.LastAddress.RowNumber, XLWorksheet.MaxNumberOfColumns);
                shiftedRangeFormula = Worksheet.Range(RangeAddress.FirstAddress, lastCell.Address);
                if (StringExtensions.IsNullOrWhiteSpace(lastCell.GetString()) && StringExtensions.IsNullOrWhiteSpace(lastCell.FormulaA1))
                    Worksheet.Internals.CellsCollection.Remove(lastCell.Address);
            }

            foreach (var ws in Worksheet.Internals.Workbook.Worksheets)
            {
                var xlWorksheet = (XLWorksheet)ws;
                foreach (var cell in xlWorksheet.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                        cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
                    else
                        cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
                }
            }

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
                var newKey = new XLAddress(c.Key.RowNumber - rowModifier, c.Key.ColumnNumber - columnModifier, false, false);
                var newCell = new XLCell(newKey, c.Value.Style, Worksheet);
                newCell.CopyValues(c.Value);
                //newCell.ShiftFormula(rowModifier * -1, columnModifier * -1);
                cellsToDelete.Add(c.Key);

                var canInsert = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ?
                    c.Key.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber :
                    c.Key.RowNumber > this.RangeAddress.LastAddress.RowNumber;

                if (canInsert)
                    cellsToInsert.Add(newKey, newCell);
            }
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
            
            List<IXLRange> mergesToRemove = new List<IXLRange>();
            foreach (var merge in Worksheet.Internals.MergedRanges)
            {
                if (this.Contains(merge))
                    mergesToRemove.Add(merge);
            }
            mergesToRemove.ForEach(r => Worksheet.Internals.MergedRanges.Remove(r));

            List<XLHyperlink> hyperlinksToRemove = new List<XLHyperlink>();
            foreach (var hl in Worksheet.Hyperlinks)
            {
                if (this.Contains(hl.Cell.AsRange()))
                    hyperlinksToRemove.Add(hl);
            }
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));

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

        public virtual IXLStyle InnerStyle
        {
            get { return this.defaultStyle; }
            set { defaultStyle = new XLStyle(this, value); }
        }

        #endregion

        public virtual IXLRange AsRange()
        {
            return Worksheet.Range(RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }

        public override string ToString()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress.ToString(), RangeAddress.LastAddress.ToString());
        }

        public  string ToStringRelative()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress.ToStringRelative(), RangeAddress.LastAddress.ToStringRelative());
        }

        public  string ToStringFixed()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress.ToStringFixed(), RangeAddress.LastAddress.ToStringFixed());
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

        public IXLRange AddToNamed(String rangeName)
        {
            return AddToNamed(rangeName, XLScope.Workbook);
        }
        public IXLRange AddToNamed(String rangeName, XLScope scope)
        {
            return AddToNamed(rangeName, scope, null);
        }
        public IXLRange AddToNamed(String rangeName, XLScope scope, String comment)
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
                namedRange.Add(this.ToStringFixed());
            }
            else
            {
                namedRanges.Add(rangeName, this.ToStringFixed(), comment);
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
                    ) || (
                           shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                        && shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                        && shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber
                        && shiftedRange.ColumnCount() > 
                            (thisRangeAddress.LastAddress.ColumnNumber - thisRangeAddress.FirstAddress.ColumnNumber + 1)
                            + (thisRangeAddress.FirstAddress.ColumnNumber - shiftedRange.RangeAddress.FirstAddress.ColumnNumber)))
                {
                    thisRangeAddress.IsInvalid = true;
                }
                else
                {
                    if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                        && shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber)
                    {
                        if (
                            (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber && columnsShifted > 0)
                            || (shiftedRange.RangeAddress.FirstAddress.ColumnNumber < thisRangeAddress.FirstAddress.ColumnNumber && columnsShifted < 0)
                            )
                            thisRangeAddress.FirstAddress = new XLAddress(
                                thisRangeAddress.FirstAddress.RowNumber, 
                                thisRangeAddress.FirstAddress.ColumnNumber + columnsShifted,
                                thisRangeAddress.FirstAddress.FixedRow, thisRangeAddress.FirstAddress.FixedColumn);

                        if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.LastAddress.ColumnNumber)
                            thisRangeAddress.LastAddress = new XLAddress(
                                thisRangeAddress.LastAddress.RowNumber, 
                                thisRangeAddress.LastAddress.ColumnNumber + columnsShifted,
                                thisRangeAddress.LastAddress.FixedRow, thisRangeAddress.LastAddress.FixedColumn);
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
                    ) || ( 
                           shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                        && shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                        && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber
                        && shiftedRange.RowCount() >
                            (thisRangeAddress.LastAddress.RowNumber - thisRangeAddress.FirstAddress.RowNumber + 1)
                            + (thisRangeAddress.FirstAddress.RowNumber - shiftedRange.RangeAddress.FirstAddress.RowNumber)))
                {
                    thisRangeAddress.IsInvalid = true;
                }
                else
                {
                    if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                        && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber)
                    {
                        if (
                            (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber && rowsShifted > 0)
                            || (shiftedRange.RangeAddress.FirstAddress.RowNumber < thisRangeAddress.FirstAddress.RowNumber && rowsShifted < 0)
                            )
                            thisRangeAddress.FirstAddress = new XLAddress(
                                thisRangeAddress.FirstAddress.RowNumber + rowsShifted,
                                thisRangeAddress.FirstAddress.ColumnNumber,
                                thisRangeAddress.FirstAddress.FixedRow, thisRangeAddress.FirstAddress.FixedColumn);

                        if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.LastAddress.RowNumber)
                            thisRangeAddress.LastAddress = new XLAddress(
                                thisRangeAddress.LastAddress.RowNumber + rowsShifted,
                                thisRangeAddress.LastAddress.ColumnNumber,
                                thisRangeAddress.LastAddress.FixedRow, thisRangeAddress.LastAddress.FixedColumn);
                    }
                }
            }
        }

        public IXLRange RangeUsed()
        {
            return this.Range(this.FirstCellUsed(), this.LastCellUsed());
        }

        public Boolean ShareString
        {
            set
            {
                Cells().ForEach(c => c.ShareString = value);
            }
        }

        public IXLHyperlinks Hyperlinks 
        {
            get
            {
                IXLHyperlinks hyperlinks = new XLHyperlinks();
                var hls = from hl in Worksheet.Hyperlinks
                          where Contains(hl.Cell.AsRange())
                          select hl;
                hls.ForEach(hl => hyperlinks.Add(hl));
                return hyperlinks;
            }
        }

        public IXLDataValidation DataValidation
        {
            get 
            {
                String address = this.RangeAddress.ToString();
                if (Worksheet.DataValidations.Where(dv => 
                    dv.Ranges.Count() == 1
                    && dv.Ranges.Contains(this.AsRange())
                    ).Any())
                {
                    return Worksheet.DataValidations.Where(dv => dv.Ranges.Contains(this.AsRange())).Single();
                }
                else
                {
                    List<IXLDataValidation> dvEmpty = new List<IXLDataValidation>();
                    foreach (var dv in Worksheet.DataValidations)
                    {
                        foreach (var dvRange in dv.Ranges)
                        {
                            if (dvRange.Intersects(this))
                            {
                                dv.Ranges.Remove(dvRange);
                                foreach (var c in dvRange.Cells())
                                {
                                    if (!this.Contains(c.Address.ToString()))
                                        dv.Ranges.Add(c.AsRange());
                                }
                                if (dv.Ranges.Count() == 0)
                                    dvEmpty.Add(dv);
                            }
                        }
                    }

                    dvEmpty.ForEach(dv => (Worksheet.DataValidations as XLDataValidations).Delete(dv));

                    var newRanges = new XLRanges(Worksheet.Internals.Workbook, Style);
                    newRanges.Add(this.AsRange());
                    var dataValidation = new XLDataValidation(newRanges, Worksheet);
                    
                    Worksheet.DataValidations.Add(dataValidation);
                    return dataValidation;
                }
            }
        }

        public Object Value
        {
            set
            {
                Cells().ForEach(c => c.Value = value);
            }
        }

        public IXLRangeBase SetValue<T>(T value)
        {
            Cells().ForEach(c => c.SetValue(value));
            return this;
        }

        public XLCellValues DataType
        {
            set
            {
                Cells().ForEach(c => c.DataType = value);
            }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges(Worksheet.Internals.Workbook, this.Style);
                retVal.Add(this.AsRange());
                return retVal;
            }
        }

        public void CopyTo(IXLRangeBase target)
        {
            CopyTo(target.FirstCell());
        }

        public void CopyTo(IXLCell target)
        {
            target.Value = this;
        }

        public void SetAutoFilter()
        {
            SetAutoFilter(true);
        }

        public void SetAutoFilter(Boolean autoFilter)
        {
            if (autoFilter)
                Worksheet.AutoFilterRange = this;
            else
                Worksheet.AutoFilterRange = null;
        }
    }
}
