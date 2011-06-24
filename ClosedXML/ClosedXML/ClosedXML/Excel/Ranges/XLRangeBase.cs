using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase : IXLRangeBase, IXLStylized
    {
        #region Fields
        protected IXLStyle m_defaultStyle;
        #endregion
        #region Constructor
        protected XLRangeBase(XLRangeAddress rangeAddress)
        {
            RangeAddress = rangeAddress;
        }
        #endregion
        #region Public properties
        public XLRangeAddress RangeAddress { get; protected set; }
        IXLRangeAddress IXLRangeBase.RangeAddress
        {
            get { return RangeAddress; }
        }
        public XLWorksheet Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }
        IXLWorksheet IXLRangeBase.Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }

        public String FormulaA1
        {
            set { Cells().ForEach(c => c.FormulaA1 = value); }
        }
        public String FormulaR1C1
        {
            set { Cells().ForEach(c => c.FormulaR1C1 = value); }
        }
        public Boolean ShareString
        {
            set { Cells().ForEach(c => c.ShareString = value); }
        }

        public IXLHyperlinks Hyperlinks
        {
            get
            {
                var hyperlinks = new XLHyperlinks();
                var hls = from hl in Worksheet.Hyperlinks
                          where Contains(hl.Cell.AsRange())
                          select hl;
                hls.ForEach(hyperlinks.Add);
                return hyperlinks;
            }
        }

        public IXLDataValidation DataValidation
        {
            get
            {
                var thisRange = AsRange();
                if (Worksheet.DataValidations.ContainsSingle(thisRange))
                {
                    return Worksheet.DataValidations.Where(dv => dv.Ranges.Contains(thisRange)).Single();
                }
                var dvEmpty = new List<IXLDataValidation>();
                foreach (var dv in Worksheet.DataValidations)
                {
                    foreach (var dvRange in dv.Ranges)
                    {
                        if (dvRange.Intersects(this))
                        {
                            dv.Ranges.Remove(dvRange);
                            foreach (var c in dvRange.Cells())
                            {
                                if (!Contains(c.Address.ToString()))
                                {
                                    dv.Ranges.Add(c.AsRange());
                                }
                            }
                            if (dv.Ranges.Count() == 0)
                            {
                                dvEmpty.Add(dv);
                            }
                        }
                    }
                }

                dvEmpty.ForEach(dv => Worksheet.DataValidations.Delete(dv));

                var newRanges = new XLRanges();
                newRanges.Add(AsRange());
                var dataValidation = new XLDataValidation(newRanges, Worksheet);

                Worksheet.DataValidations.Add(dataValidation);
                return dataValidation;
            }
        }

        public Object Value
        {
            set { Cells().ForEach(c => c.Value = value); }
        }

        public XLCellValues DataType
        {
            set { Cells().ForEach(c => c.DataType = value); }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                retVal.Add(AsRange());
                return retVal;
            }
        }
        #endregion
        IXLCell IXLRangeBase.FirstCell()
        {
            return FirstCell();
        }
        public XLCell FirstCell()
        {
            return Cell(1, 1);
        }
        IXLCell IXLRangeBase.LastCell()
        {
            return LastCell();
        }
        public XLCell LastCell()
        {
            return Cell(RowCount(), ColumnCount());
        }

        IXLCell IXLRangeBase.FirstCellUsed()
        {
            return FirstCellUsed(false);
        }
        IXLCell IXLRangeBase.FirstCellUsed(bool includeStyles)
        {
            return FirstCellUsed(includeStyles);
        }
        public XLCell FirstCellUsed()
        {
            return FirstCellUsed(false);
        }
        public XLCell FirstCellUsed(Boolean includeStyles)
        {
            var cellsUsed = CellsUsed(includeStyles);

            if (!cellsUsed.Any<XLCell>())
            {
                return null;
            }
            var firstRow = cellsUsed.Min<XLCell>(c => c.Address.RowNumber);
            var firstColumn = cellsUsed.Min<XLCell>(c => c.Address.ColumnNumber);
            var mergedRanges = Worksheet.Internals.MergedRanges.GetContainingMergedRanges(GetSheetRange());
            foreach (var mrange in mergedRanges)
            {
                firstRow = Math.Max(mrange.FirstAddress.RowNumber, firstRow);
                firstColumn = Math.Max(mrange.FirstAddress.ColumnNumber, firstColumn);
            }
            return Worksheet.Cell(firstRow, firstColumn);
        }

        IXLCell IXLRangeBase.LastCellUsed()
        {
            return LastCellUsed(false);
        }
        IXLCell IXLRangeBase.LastCellUsed(bool includeStyles)
        {
            return LastCellUsed(includeStyles);
        }
        public XLCell LastCellUsed()
        {
            return LastCellUsed(false);
        }
        public XLCell LastCellUsed(bool includeStyles)
        {
            var cellsUsed = CellsUsed(includeStyles);
            if (!cellsUsed.Any<XLCell>())
            {
                return null;
            }

            var lastRow = cellsUsed.Max<XLCell>(c => c.Address.RowNumber);
            var lastColumn = cellsUsed.Max<XLCell>(c => c.Address.ColumnNumber);
            var mergedRanges = Worksheet.Internals.MergedRanges.GetContainingMergedRanges(GetSheetRange());
            foreach (var mrange in mergedRanges)
            {
                lastRow = Math.Max(mrange.LastAddress.RowNumber, lastRow);
                lastColumn = Math.Max(mrange.LastAddress.ColumnNumber, lastColumn);
            }
            return Worksheet.Cell(lastRow, lastColumn);
            
        }

        public XLCell Cell(Int32 row, Int32 column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public XLCell Cell(String cellAddressInRange)
        {
            return Cell(XLAddress.Create(Worksheet, cellAddressInRange));
        }

        public XLCell Cell(Int32 row, String column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }
        public XLCell Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange.RowNumber, cellAddressInRange.ColumnNumber);
        }

        public XLCell Cell(XLAddress cellAddressInRange)
        {
            var absoluteAddress = cellAddressInRange + RangeAddress.FirstAddress - 1;
            if (Worksheet.Internals.CellsCollection.ContainsKey(absoluteAddress))
            {
                return Worksheet.Internals.CellsCollection[absoluteAddress];
            }
            IXLStyle style = Style;
            if (Style != null && Style.Equals(Worksheet.Style))
            {
                if (Worksheet.Internals.RowsCollection.ContainsKey(absoluteAddress.RowNumber)
                    && !Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style.Equals(Worksheet.Style))
                {
                    style = Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style;
                }
                else if (Worksheet.Internals.ColumnsCollection.ContainsKey(absoluteAddress.ColumnNumber)
                         && !Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style.Equals(Worksheet.Style))
                {
                    style = Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style;
                }
            }
            var newCell = new XLCell(Worksheet, absoluteAddress, style);
            Worksheet.Internals.CellsCollection.Add(absoluteAddress, newCell);
            return newCell;
        }

        public Int32 RowCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }
        public Int32 RowNumber()
        {
            return RangeAddress.FirstAddress.RowNumber;
        }
        public Int32 ColumnCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }
        public Int32 ColumnNumber()
        {
            return RangeAddress.FirstAddress.ColumnNumber;
        }
        public String ColumnLetter()
        {
            return RangeAddress.FirstAddress.ColumnLetter;
        }

        public virtual XLRange Range(String rangeAddressStr)
        {
            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressStr);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLCell firstCell, IXLCell lastCell)
        {
            return Range(firstCell.Address, lastCell.Address);
        }
        public XLRange Range(String firstCellAddress, String lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(XLAddress.Create(Worksheet, firstCellAddress), XLAddress.Create(Worksheet, lastCellAddress));
            return Range(rangeAddress);
        }
        public XLRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            var rangeAddress = new XLRangeAddress(new XLAddress(Worksheet, firstCellRow, firstCellColumn, false, false),
                                                  new XLAddress(Worksheet, lastCellRow, lastCellColumn, false, false));
            return Range(rangeAddress);
        }
        public XLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(firstCellAddress as XLAddress, lastCellAddress as XLAddress);
            return Range(rangeAddress);
        }
        public XLRange Range(IXLRangeAddress rangeAddress)
        {
            var newFirstCellAddress = (XLAddress) rangeAddress.FirstAddress + RangeAddress.FirstAddress - 1;
            newFirstCellAddress.FixedRow = rangeAddress.FirstAddress.FixedRow;
            newFirstCellAddress.FixedColumn = rangeAddress.FirstAddress.FixedColumn;

            var newLastCellAddress = (XLAddress) rangeAddress.LastAddress + RangeAddress.FirstAddress - 1;
            newLastCellAddress.FixedRow = rangeAddress.LastAddress.FixedRow;
            newLastCellAddress.FixedColumn = rangeAddress.LastAddress.FixedColumn;

            var newRangeAddress = new XLRangeAddress(newFirstCellAddress, newLastCellAddress);
            var xlRangeParameters = new XLRangeParameters(newRangeAddress, Style);
            if (
                    newFirstCellAddress.RowNumber < RangeAddress.FirstAddress.RowNumber
                    || newFirstCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                    || newLastCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                    || newFirstCellAddress.ColumnNumber < RangeAddress.FirstAddress.ColumnNumber
                    || newFirstCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                    || newLastCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                    )
            {
                throw new ArgumentOutOfRangeException(String.Format("The cells {0} and {1} are outside the range '{2}'.",
                                                                    newFirstCellAddress,
                                                                    newLastCellAddress,
                                                                    ToString()));
            }

            return new XLRange(xlRangeParameters);
        }

        public IXLRanges Ranges(String ranges)
        {
            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()));
            }
            return retVal;
        }
        public IXLRanges Ranges(params String[] ranges)
        {
            var retVal = new XLRanges();
            foreach (var pair in ranges)
            {
                retVal.Add(Range(pair));
            }
            return retVal;
        }

        protected String FixColumnAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
            {
                return "A" + address;
            }
            return address;
        }
        protected String FixRowAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
            {
                return ExcelHelper.GetColumnLetterFromNumber(test) + "1";
            }
            return address;
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, false, false);
            cells.Add(RangeAddress);
            return cells;
        }
        public IXLCells Cells(String cells)
        {
            return Ranges(cells).Cells();
        }
        public IXLCells CellsUsed()
        {
            var cells = new XLCells(false, true, false);
            cells.Add(RangeAddress);
            return cells;
        }
        IXLCells IXLRangeBase.CellsUsed(Boolean includeStyles)
        {
            return CellsUsed(includeStyles);
        }
        public XLCells CellsUsed(bool includeStyles)
        {
            var cells = new XLCells(false, true, includeStyles);
            cells.Add(RangeAddress);
            return cells;
        }

        public IXLRange Merge()
        {
            Boolean foundOne = false;
            foreach (var m in (Worksheet).Internals.MergedRanges)
            {
                
                if (m.Equals(RangeAddress))
                {
                    foundOne = true;
                    break;
                }
            }

            if (!foundOne)
            {
                Worksheet.Internals.MergedRanges.Add(GetSheetRange());
            }
            return AsRange();
        }
        public IXLRange Unmerge()
        {
            foreach (var m in (Worksheet).Internals.MergedRanges)
            {
                if (m.Equals(RangeAddress))
                {
                    Worksheet.Internals.MergedRanges.Remove(m);
                    break;
                }
            }
            return AsRange();
        }

        public virtual SheetRange GetSheetRange()
        {
            return RangeAddress.GetSheetRange();
        }

        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            return InsertColumnsAfter(numberOfColumns, true);
        }
        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = InsertColumnsAfter(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet,
                                      RangeAddress.FirstAddress.RowNumber,
                                      RangeAddress.FirstAddress.ColumnNumber,
                                      RangeAddress.FirstAddress.FixedRow,
                                      RangeAddress.FirstAddress.FixedColumn),
                        new XLAddress(Worksheet,
                                      RangeAddress.LastAddress.RowNumber,
                                      RangeAddress.LastAddress.ColumnNumber + numberOfColumns,
                                      RangeAddress.LastAddress.FixedRow,
                                      RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeColumns InsertColumnsAfter(Boolean onlyUsedCells, Int32 numberOfColumns)
        {
            var columnCount = ColumnCount();
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber + columnCount;
            if (firstColumn > ExcelHelper.MaxColumnNumber)
            {
                firstColumn = ExcelHelper.MaxColumnNumber;
            }
            var lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > ExcelHelper.MaxColumnNumber)
            {
                lastColumn = ExcelHelper.MaxColumnNumber;
            }

            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var lastRow = firstRow + RowCount() - 1;
            if (lastRow > ExcelHelper.MaxRowNumber)
            {
                lastRow = ExcelHelper.MaxRowNumber;
            }

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertColumnsBefore(onlyUsedCells, numberOfColumns);
        }
        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            return InsertColumnsBefore(numberOfColumns, false);
        }
        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = InsertColumnsBefore(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet,
                                      RangeAddress.FirstAddress.RowNumber,
                                      RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                                      RangeAddress.FirstAddress.FixedRow,
                                      RangeAddress.FirstAddress.FixedColumn),
                        new XLAddress(Worksheet,
                                      RangeAddress.LastAddress.RowNumber,
                                      RangeAddress.LastAddress.ColumnNumber,
                                      RangeAddress.LastAddress.FixedRow,
                                      RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeColumns InsertColumnsBefore(Boolean onlyUsedCells, Int32 numberOfColumns)
        {
            foreach (var ws in (Worksheet).Internals.Workbook.WorksheetsInternal)
            {
                foreach (var cell in ws.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaColumns((XLRange) AsRange(), numberOfColumns);
                }
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var cellsToBlank = new List<IXLAddress>();
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var lastRow = RangeAddress.FirstAddress.RowNumber + RowCount() - 1;

            if (!onlyUsedCells)
            {
                var lastColumn = Worksheet.LastColumnUsed().ColumnNumber();

                for (var co = lastColumn; co >= firstColumn; co--)
                {
                    for (var ro = lastRow; ro >= firstRow; ro--)
                    {
                        var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                        var newColumn = co + numberOfColumns;
                        var newKey = new XLAddress(Worksheet, ro, newColumn, false, false);
                        IXLCell oldCell;
                        if ((Worksheet).Internals.CellsCollection.ContainsKey(oldKey))
                        {
                            oldCell = (Worksheet).Internals.CellsCollection[oldKey];
                        }
                        else
                        {
                            oldCell = Worksheet.Cell(oldKey);
                        }
                        var newCell = new XLCell(Worksheet, newKey, oldCell.Style);
                        newCell.CopyValues((XLCell) oldCell);
                        cellsToInsert.Add(newKey, newCell);
                        cellsToDelete.Add(oldKey);
                        if (oldKey.ColumnNumber < firstColumn + numberOfColumns)
                        {
                            cellsToBlank.Add(oldKey);
                        }
                    }
                }
            }
            else
            {
                foreach (var c in (Worksheet).Internals.CellsCollection
                        .Where(c =>
                               c.Key.ColumnNumber >= firstColumn
                               && c.Key.RowNumber >= firstRow
                               && c.Key.RowNumber <= lastRow
                        ))
                {
                    var newColumn = c.Key.ColumnNumber + numberOfColumns;
                    var newKey = new XLAddress(Worksheet, c.Key.RowNumber, newColumn, false, false);
                    var newCell = new XLCell(Worksheet, newKey, c.Value.Style);
                    newCell.CopyValues(c.Value);
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Key);
                    if (c.Key.ColumnNumber < firstColumn + numberOfColumns)
                    {
                        cellsToBlank.Add(c.Key);
                    }
                }
            }
            cellsToDelete.ForEach(c => (Worksheet).Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => (Worksheet).Internals.CellsCollection.Add(c.Key, c.Value));
            foreach (var c in cellsToBlank)
            {
                IXLStyle styleToUse;
                if ((Worksheet).Internals.RowsCollection.ContainsKey(c.RowNumber))
                {
                    styleToUse = (Worksheet).Internals.RowsCollection[c.RowNumber].Style;
                }
                else
                {
                    styleToUse = Worksheet.Style;
                }
                Worksheet.Cell(c.RowNumber, c.ColumnNumber).Style = styleToUse;
            }

            (Worksheet).NotifyRangeShiftedColumns((XLRange) AsRange(), numberOfColumns);
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
            var retVal = InsertRowsBelow(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet,
                                      RangeAddress.FirstAddress.RowNumber,
                                      RangeAddress.FirstAddress.ColumnNumber,
                                      RangeAddress.FirstAddress.FixedRow,
                                      RangeAddress.FirstAddress.FixedColumn),
                        new XLAddress(Worksheet,
                                      RangeAddress.LastAddress.RowNumber + numberOfRows,
                                      RangeAddress.LastAddress.ColumnNumber,
                                      RangeAddress.LastAddress.FixedRow,
                                      RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeRows InsertRowsBelow(Boolean onlyUsedCells, Int32 numberOfRows)
        {
            var rowCount = RowCount();
            var firstRow = RangeAddress.FirstAddress.RowNumber + rowCount;
            if (firstRow > ExcelHelper.MaxRowNumber)
            {
                firstRow = ExcelHelper.MaxRowNumber;
            }
            var lastRow = firstRow + RowCount() - 1;
            if (lastRow > ExcelHelper.MaxRowNumber)
            {
                lastRow = ExcelHelper.MaxRowNumber;
            }

            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > ExcelHelper.MaxColumnNumber)
            {
                lastColumn = ExcelHelper.MaxColumnNumber;
            }

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertRowsAbove(onlyUsedCells, numberOfRows);
        }

        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows)
        {
            return InsertRowsAbove(numberOfRows, false);
        }
        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows, Boolean expandRange)
        {
            var retVal = InsertRowsAbove(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet,
                                      RangeAddress.FirstAddress.RowNumber - numberOfRows,
                                      RangeAddress.FirstAddress.ColumnNumber,
                                      RangeAddress.FirstAddress.FixedRow,
                                      RangeAddress.FirstAddress.FixedColumn),
                        new XLAddress(Worksheet,
                                      RangeAddress.LastAddress.RowNumber,
                                      RangeAddress.LastAddress.ColumnNumber,
                                      RangeAddress.LastAddress.FixedRow,
                                      RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }
        public IXLRangeRows InsertRowsAbove(Boolean onlyUsedCells, Int32 numberOfRows)
        {
            foreach (var ws in (Worksheet).Internals.Workbook.WorksheetsInternal)
            {
                foreach (var cell in ws.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaRows((XLRange) AsRange(), numberOfRows);
                }
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var cellsToBlank = new List<IXLAddress>();
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;

            if (!onlyUsedCells)
            {
                var lastRow = Worksheet.LastRowUsed().RowNumber();

                for (var ro = lastRow; ro >= firstRow; ro--)
                {
                    for (var co = lastColumn; co >= firstColumn; co--)
                    {
                        var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                        var newRow = ro + numberOfRows;
                        var newKey = new XLAddress(Worksheet, newRow, co, false, false);
                        XLCell oldCell;
                        if ((Worksheet).Internals.CellsCollection.ContainsKey(oldKey))
                        {
                            oldCell = (Worksheet).Internals.CellsCollection[oldKey];
                        }
                        else
                        {
                            oldCell = Worksheet.Cell(oldKey);
                        }
                        var newCell = new XLCell(Worksheet, newKey, oldCell.Style);
                        newCell.CopyFrom(oldCell);
                        cellsToInsert.Add(newKey, newCell);
                        cellsToDelete.Add(oldKey);
                        if (oldKey.RowNumber < firstRow + numberOfRows)
                        {
                            cellsToBlank.Add(oldKey);
                        }
                    }
                }
            }
            else
            {
                foreach (var c in (Worksheet).Internals.CellsCollection
                        .Where(c =>
                               c.Key.RowNumber >= firstRow
                               && c.Key.ColumnNumber >= firstColumn
                               && c.Key.ColumnNumber <= lastColumn
                        ))
                {
                    var newRow = c.Key.RowNumber + numberOfRows;
                    var newKey = new XLAddress(Worksheet, newRow, c.Key.ColumnNumber, false, false);
                    var newCell = new XLCell(Worksheet, newKey, c.Value.Style);
                    newCell.CopyFrom(c.Value);
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Key);
                    if (c.Key.RowNumber < firstRow + numberOfRows)
                    {
                        cellsToBlank.Add(c.Key);
                    }
                }
            }
            cellsToDelete.ForEach(c => (Worksheet).Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => (Worksheet).Internals.CellsCollection.Add(c.Key, c.Value));
            foreach (var c in cellsToBlank)
            {
                IXLStyle styleToUse;
                if ((Worksheet).Internals.ColumnsCollection.ContainsKey(c.ColumnNumber))
                {
                    styleToUse = (Worksheet).Internals.ColumnsCollection[c.ColumnNumber].Style;
                }
                else
                {
                    styleToUse = Worksheet.Style;
                }
                Worksheet.Cell(c.RowNumber, c.ColumnNumber).Style = styleToUse;
            }
            (Worksheet).NotifyRangeShiftedRows((XLRange) AsRange(), numberOfRows);
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
            (Worksheet).Internals.CellsCollection.RemoveAll(c =>
                                                            c.Address.ColumnNumber >= RangeAddress.FirstAddress.ColumnNumber
                                                            && c.Address.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber
                                                            && c.Address.RowNumber >= RangeAddress.FirstAddress.RowNumber
                                                            && c.Address.RowNumber <= RangeAddress.LastAddress.RowNumber
                    );

            ClearMerged();

            List<XLHyperlink> hyperlinksToRemove = new List<XLHyperlink>();
            foreach (var hl in Worksheet.Hyperlinks)
            {
                if (Contains(hl.Cell.AsRange()))
                {
                    hyperlinksToRemove.Add(hl);
                }
            }
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));
        }

        public void ClearStyles()
        {
            foreach (var cell in CellsUsed(true))
            {
                var newStyle = new XLStyle((XLCell) cell, Worksheet.Style);
                newStyle.NumberFormat = cell.Style.NumberFormat;
                cell.Style = newStyle;
            }
        }

        private void ClearMerged()
        {
            var mergeToDelete = new List<SheetRange>();
            foreach (var merge in (Worksheet).Internals.MergedRanges)
            {
                if (Intersects(merge))
                {
                    mergeToDelete.Add(merge);
                }
            }
            mergeToDelete.ForEach(m => (Worksheet).Internals.MergedRanges.Remove(m));
        }

        public bool Contains(String rangeAddress)
        {
            String addressToUse;
            if (rangeAddress.Contains("!"))
            {
                addressToUse = rangeAddress.Substring(rangeAddress.IndexOf("!") + 1);
            }
            else
            {
                addressToUse = rangeAddress;
            }

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                String[] arrRange = addressToUse.Split(':');
                firstAddress = XLAddress.Create(Worksheet, arrRange[0]);
                lastAddress = XLAddress.Create(Worksheet, arrRange[1]);
            }
            else
            {
                firstAddress = XLAddress.Create(Worksheet, addressToUse);
                lastAddress = XLAddress.Create(Worksheet, addressToUse);
            }
            return Contains(firstAddress, lastAddress);
        }
        public bool Contains(IXLRangeBase range)
        {
            return Contains((XLAddress) range.RangeAddress.FirstAddress, (XLAddress) range.RangeAddress.LastAddress);
        }
        public bool Contains(XLAddress first, XLAddress last)
        {
            return Contains(first) && Contains(last);
        }
        public bool Contains(XLAddress address)
        {
            return RangeAddress.FirstAddress.RowNumber <= address.RowNumber && address.RowNumber <= RangeAddress.LastAddress.RowNumber &&
                   RangeAddress.FirstAddress.ColumnNumber <= address.ColumnNumber && address.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber;
        }
        public bool Contains(SheetRange range)
        {
            return Contains(range.FirstAddress, range.LastAddress);
        }
        public bool Contains(SheetPoint first, SheetPoint last)
        {
            return Contains(first) && Contains(last);
        }
        public bool Contains(SheetPoint point)
        {
            return RangeAddress.FirstAddress.RowNumber <= point.RowNumber && point.RowNumber <= RangeAddress.LastAddress.RowNumber &&
                   RangeAddress.FirstAddress.ColumnNumber <= point.ColumnNumber && point.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber;
        }

        public bool Intersects(string rangeAddress)
        {
            return Intersects(Range(rangeAddress));
        }
        public bool Intersects(IXLRangeBase range)
        {
            if (range.RangeAddress.IsInvalid || RangeAddress.IsInvalid)
            {
                return false;
            }
            var ma = range.RangeAddress;
            var ra = RangeAddress;

            return !( // See if the two ranges intersect...
                    ma.FirstAddress.ColumnNumber > ra.LastAddress.ColumnNumber
                    || ma.LastAddress.ColumnNumber < ra.FirstAddress.ColumnNumber
                    || ma.FirstAddress.RowNumber > ra.LastAddress.RowNumber
                    || ma.LastAddress.RowNumber < ra.FirstAddress.RowNumber
                    );
        }
        public bool Intersects(SheetRange range)
        {
            if (RangeAddress.IsInvalid)
            {
                return false;
            }
            var ra = RangeAddress;

            return !( // See if the two ranges intersect...
                    range.FirstAddress.ColumnNumber > ra.LastAddress.ColumnNumber
                    || range.LastAddress.ColumnNumber < ra.FirstAddress.ColumnNumber
                    || range.FirstAddress.RowNumber > ra.LastAddress.RowNumber
                    || range.LastAddress.RowNumber < ra.FirstAddress.RowNumber
                    );
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            var numberOfRows = RowCount();
            var numberOfColumns = ColumnCount();
            IXLRange shiftedRangeFormula;
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                var lastCell = Worksheet.Cell(ExcelHelper.MaxRowNumber, RangeAddress.LastAddress.ColumnNumber);
                shiftedRangeFormula = Worksheet.Range(RangeAddress.FirstAddress, lastCell.Address);
                if (StringExtensions.IsNullOrWhiteSpace(lastCell.GetString()) && StringExtensions.IsNullOrWhiteSpace(lastCell.FormulaA1))
                {
                    (Worksheet).Internals.CellsCollection.Remove(lastCell.Address);
                }
            }
            else
            {
                var lastCell = Worksheet.Cell(RangeAddress.LastAddress.RowNumber, ExcelHelper.MaxColumnNumber);
                shiftedRangeFormula = Worksheet.Range(RangeAddress.FirstAddress, lastCell.Address);
                if (StringExtensions.IsNullOrWhiteSpace(lastCell.GetString()) && StringExtensions.IsNullOrWhiteSpace(lastCell.FormulaA1))
                {
                    (Worksheet).Internals.CellsCollection.Remove(lastCell.Address);
                }
            }

            foreach (var ws in (Worksheet).Internals.Workbook.WorksheetsInternal)
            {
                foreach (var cell in ws.Internals.CellsCollection.Values.Where(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                    {
                        cell.ShiftFormulaRows((XLRange) shiftedRangeFormula, numberOfRows*-1);
                    }
                    else
                    {
                        cell.ShiftFormulaColumns((XLRange) shiftedRangeFormula, numberOfColumns*-1);
                    }
                }
            }

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var shiftLeftQuery = (Worksheet).Internals.CellsCollection
                    .Where(c =>
                           c.Key.RowNumber >= RangeAddress.FirstAddress.RowNumber
                           && c.Key.RowNumber <= RangeAddress.LastAddress.RowNumber
                           && c.Key.ColumnNumber >= RangeAddress.FirstAddress.ColumnNumber);

            var shiftUpQuery = (Worksheet).Internals.CellsCollection
                    .Where(c =>
                           c.Key.ColumnNumber >= RangeAddress.FirstAddress.ColumnNumber
                           && c.Key.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber
                           && c.Key.RowNumber >= RangeAddress.FirstAddress.RowNumber);

            var columnModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? ColumnCount() : 0;
            var rowModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp ? RowCount() : 0;
            var cellsQuery = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? shiftLeftQuery : shiftUpQuery;
            foreach (var c in cellsQuery)
            {
                var newKey = new XLAddress(Worksheet, c.Key.RowNumber - rowModifier, c.Key.ColumnNumber - columnModifier, false, false);
                var newCell = new XLCell(Worksheet, newKey, c.Value.Style);
                newCell.CopyValues(c.Value);
                //newCell.ShiftFormula(rowModifier * -1, columnModifier * -1);
                cellsToDelete.Add(c.Key);

                var canInsert = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft
                                        ? c.Key.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                                        : c.Key.RowNumber > RangeAddress.LastAddress.RowNumber;

                if (canInsert)
                {
                    cellsToInsert.Add(newKey, newCell);
                }
            }
            cellsToDelete.ForEach(c => (Worksheet).Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => (Worksheet).Internals.CellsCollection.Add(c.Key, c.Value));

            var mergesToRemove = new List<SheetRange>();
            foreach (var merge in (Worksheet).Internals.MergedRanges)
            {
                if (Contains(merge))
                {
                    mergesToRemove.Add(merge);
                }
            }
            mergesToRemove.ForEach(r => (Worksheet).Internals.MergedRanges.Remove(r));

            List<XLHyperlink> hyperlinksToRemove = new List<XLHyperlink>();
            foreach (var hl in Worksheet.Hyperlinks)
            {
                if (Contains(hl.Cell.AsRange()))
                {
                    hyperlinksToRemove.Add(hl);
                }
            }
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));

            var shiftedRange = (XLRange) AsRange();
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                (Worksheet).NotifyRangeShiftedRows(shiftedRange, rowModifier*-1);
            }
            else
            {
                (Worksheet).NotifyRangeShiftedColumns(shiftedRange, columnModifier*-1);
            }
        }
        #region IXLStylized Members
        public virtual IXLStyle Style
        {
            get { return m_defaultStyle; }
            set { Cells().ForEach(c => c.Style = value); }
        }

        public virtual IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                foreach (var cell in Cells())
                {
                    yield return cell.Style;
                }
                UpdatingStyle = false;
            }
        }

        public virtual Boolean UpdatingStyle { get; set; }

        public virtual IXLStyle InnerStyle
        {
            get { return m_defaultStyle; }
            set { m_defaultStyle = new XLStyle(this, value); }
        }
        #endregion
        public virtual IXLRange AsRange()
        {
            return Worksheet.Range(RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }

        public override string ToString()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }

        public string ToStringRelative()
        {
            return String.Format("'{0}'!{1}:{2}",
                                 Worksheet.Name,
                                 RangeAddress.FirstAddress.ToStringRelative(),
                                 RangeAddress.LastAddress.ToStringRelative());
        }

        public string ToStringFixed()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress.ToStringFixed(), RangeAddress.LastAddress.ToStringFixed());
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
                namedRanges = (Worksheet).Internals.Workbook.NamedRanges;
            }
            else
            {
                namedRanges = Worksheet.NamedRanges;
            }

            if (namedRanges.Any(nr => nr.Name.ToLower() == rangeName.ToLower()))
            {
                var namedRange = namedRanges.Where(nr => nr.Name.ToLower() == rangeName.ToLower()).Single();
                namedRange.Add((Worksheet).Internals.Workbook, ToStringFixed());
            }
            else
            {
                namedRanges.Add(rangeName, ToStringFixed(), comment);
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
                                (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber &&
                                 columnsShifted > 0)
                                ||
                                (shiftedRange.RangeAddress.FirstAddress.ColumnNumber < thisRangeAddress.FirstAddress.ColumnNumber &&
                                 columnsShifted < 0)
                                )
                        {
                            thisRangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                                          thisRangeAddress.FirstAddress.RowNumber,
                                                                          thisRangeAddress.FirstAddress.ColumnNumber + columnsShifted,
                                                                          thisRangeAddress.FirstAddress.FixedRow,
                                                                          thisRangeAddress.FirstAddress.FixedColumn);
                        }

                        if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.LastAddress.ColumnNumber)
                        {
                            thisRangeAddress.LastAddress = new XLAddress(Worksheet,
                                                                         thisRangeAddress.LastAddress.RowNumber,
                                                                         thisRangeAddress.LastAddress.ColumnNumber + columnsShifted,
                                                                         thisRangeAddress.LastAddress.FixedRow,
                                                                         thisRangeAddress.LastAddress.FixedColumn);
                        }
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
                        {
                            thisRangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                                          thisRangeAddress.FirstAddress.RowNumber + rowsShifted,
                                                                          thisRangeAddress.FirstAddress.ColumnNumber,
                                                                          thisRangeAddress.FirstAddress.FixedRow,
                                                                          thisRangeAddress.FirstAddress.FixedColumn);
                        }

                        if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.LastAddress.RowNumber)
                        {
                            thisRangeAddress.LastAddress = new XLAddress(Worksheet,
                                                                         thisRangeAddress.LastAddress.RowNumber + rowsShifted,
                                                                         thisRangeAddress.LastAddress.ColumnNumber,
                                                                         thisRangeAddress.LastAddress.FixedRow,
                                                                         thisRangeAddress.LastAddress.FixedColumn);
                        }
                    }
                }
            }
        }

        public IXLRange RangeUsed()
        {
            return RangeUsed(false);
        }

        public IXLRange RangeUsed(bool includeStyles)
        {
            var firstCell = FirstCellUsed(includeStyles);
            if (firstCell == null)
            {
                return null;
            }
            var lastCell = LastCellUsed(includeStyles);
            return Worksheet.Range(firstCell, lastCell);
        }

        public IXLRangeBase SetValue<T>(T value)
        {
            Cells().ForEach(c => c.SetValue(value));
            return this;
        }

        public virtual void CopyTo(IXLRangeBase target)
        {
            CopyTo(target.FirstCell());
        }

        public virtual void CopyTo(IXLCell target)
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
            {
                Worksheet.AutoFilterRange = this;
            }
            else
            {
                Worksheet.AutoFilterRange = null;
            }
        }

        //public IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn)
        //{
        //    IXLChart chart = new XLChart(Worksheet);
        //    chart.FirstRow = firstRow;
        //    chart.LastRow = lastRow;
        //    chart.LastColumn = lastColumn;
        //    chart.FirstColumn = firstColumn;
        //    Worksheet.Charts.Add(chart);
        //    return chart;
        //}
    }
}