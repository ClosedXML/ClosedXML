// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {
        public Int32 MaxColumnUsed;
        public Int32 MaxRowUsed;
        public Dictionary<Int32, Int32> RowsUsed = new Dictionary<int, int>();

        public XLCellsCollection()
        {
            Clear();
        }

        public Int32 Count { get; private set; }
        internal Dictionary<Int32, Int32> ColumnsUsed { get; } = new Dictionary<int, int>();
        internal Dictionary<Int32, HashSet<Int32>> Deleted { get; } = new Dictionary<int, HashSet<int>>();
        internal Dictionary<int, Dictionary<int, XLCell>> RowsCollection { get; } = new Dictionary<int, Dictionary<int, XLCell>>();

        public void Add(XLSheetPoint sheetPoint, XLCell cell)
        {
            Add(sheetPoint.Row, sheetPoint.Column, cell);
        }

        public void Add(Int32 row, Int32 column, XLCell cell)
        {
            Count++;

            IncrementUsage(RowsUsed, row);
            IncrementUsage(ColumnsUsed, column);

            if (!RowsCollection.TryGetValue(row, out Dictionary<int, XLCell> columnsCollection))
            {
                columnsCollection = new Dictionary<int, XLCell>();
                RowsCollection.Add(row, columnsCollection);
            }
            columnsCollection.Add(column, cell);
            if (row > MaxRowUsed) MaxRowUsed = row;
            if (column > MaxColumnUsed) MaxColumnUsed = column;

            if (Deleted.TryGetValue(row, out HashSet<int> delHash))
                delHash.Remove(column);
        }

        public void Clear()
        {
            Count = 0;
            RowsUsed.Clear();
            ColumnsUsed.Clear();

            RowsCollection.Clear();
            MaxRowUsed = 0;
            MaxColumnUsed = 0;
        }

        public Boolean Contains(Int32 row, Int32 column)
        {
            return RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                && columnsCollection.ContainsKey(column);
        }

        public int FirstColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            int firstColumnUsed = finalColumn;
            var found = false;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = columnStart; co <= firstColumnUsed; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell))
                            && co <= firstColumnUsed)
                        {
                            firstColumnUsed = co;
                            found = true;
                            break;
                        }
                    }
                }
            }

            return found ? firstColumnUsed : 0;
        }

        public int FirstRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            return ro;
                    }
                }
            }

            return 0;
        }

        public XLCell GetCell(Int32 row, Int32 column)
        {
            if (row > MaxRowUsed || column > MaxColumnUsed)
                return null;

            if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection))
            {
                return columnsCollection.TryGetValue(column, out XLCell cell) ? cell : null;
            }
            return null;
        }

        public XLCell GetCell(XLSheetPoint sp)
        {
            return GetCell(sp.Row, sp.Column);
        }

        public IEnumerable<XLCell> GetCellsInColumn(Int32 column)
        {
            return GetCells(1, column, MaxRowUsed, column);
        }

        public IEnumerable<XLCell> GetCellsInRow(Int32 row)
        {
            return GetCells(row, 1, row, MaxColumnUsed);
        }

        public IEnumerable<XLSheetPoint> GetSheetPoints(Int32 rowStart, Int32 columnStart,
                                                        Int32 rowEnd, Int32 columnEnd)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.ContainsKey(co))
                            yield return new XLSheetPoint(ro, co);
                    }
                }
            }
        }

        public int LastColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, Boolean> predicate = null)
        {
            int maxCo = 0;
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = finalRow; ro >= rowStart; ro--)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<int, XLCell> columnsCollection))
                {
                    for (int co = finalColumn; co >= columnStart && co > maxCo; co--)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            maxCo = co;
                    }
                }
            }
            return maxCo;
        }

        public int LastRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = finalRow; ro >= rowStart; ro--)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = finalColumn; co >= columnStart; co--)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            return ro;
                    }
                }
            }
            return 0;
        }

        public Int32 MaxColumnInRow(Int32 row)
        {
            if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                && columnsCollection.Any())

                return columnsCollection.Keys.Max();

            return 0;
        }

        public Int32 MaxRowInColumn(Int32 column)
        {
            for (int row = MaxRowUsed; row >= 1; row--)
            {
                if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                    && columnsCollection.ContainsKey(column))

                    return row;
            }

            return 0;
        }

        public Int32 MinColumnInRow(Int32 row)
        {
            if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                && columnsCollection.Any())

                return columnsCollection.Keys.Min();

            return 0;
        }

        public Int32 MinRowInColumn(Int32 column)
        {
            for (int row = 1; row <= MaxRowUsed; row++)
            {
                if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                    && columnsCollection.ContainsKey(column))

                    return row;
            }

            return 0;
        }

        public void Remove(XLSheetPoint sheetPoint)
        {
            Remove(sheetPoint.Row, sheetPoint.Column);
        }

        public void Remove(Int32 row, Int32 column)
        {
            Count--;
            var rowRemoved = DecrementUsage(RowsUsed, row);
            var columnRemoved = DecrementUsage(ColumnsUsed, column);

            if (rowRemoved && row == MaxRowUsed)
            {
                MaxRowUsed = RowsUsed.Keys.Any()
                    ? RowsUsed.Keys.Max()
                    : 0;
            }

            if (columnRemoved && column == MaxColumnUsed)
            {
                MaxColumnUsed = ColumnsUsed.Keys.Any()
                    ? ColumnsUsed.Keys.Max()
                    : 0;
            }

            if (Deleted.TryGetValue(row, out HashSet<Int32> delHash))
            {
                if (!delHash.Contains(column))
                    delHash.Add(column);
            }
            else
            {
                delHash = new HashSet<int>();
                delHash.Add(column);
                Deleted.Add(row, delHash);
            }

            if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection))
            {
                columnsCollection.Remove(column);
                if (columnsCollection.Count == 0)
                {
                    RowsCollection.Remove(row);
                }
            }
        }

        public void RemoveAll(Int32 rowStart, Int32 columnStart,
                              Int32 rowEnd, Int32 columnEnd)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<int, XLCell> columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.ContainsKey(co))
                            Remove(ro, co);
                    }
                }
            }
        }

        internal void ChangeColumnNumbers(IDictionary<int, int> columnNumberMappings, XLRangeAddress affectedRange)
        {
            ChangeColumnNumbers(columnNumberMappings, this.RowsCollection, affectedRange);
        }

        internal void ChangeRowNumbers(IDictionary<int, int> rowNumberMappings, XLRangeAddress affectedRange)
        {
            ChangeRowNumbers(rowNumberMappings, this.RowsCollection, affectedRange);
        }

        internal IEnumerable<XLCell> GetCells(Int32 rowStart, Int32 columnStart,
                                            Int32 rowEnd, Int32 columnEnd,
                                            Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }
        }

        internal IEnumerable<XLCell> GetCells()
        {
            return GetCells(1, 1, MaxRowUsed, MaxColumnUsed);
        }

        internal IEnumerable<XLCell> GetCells(Func<IXLCell, Boolean> predicate)
        {
            for (int ro = 1; ro <= MaxRowUsed; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out Dictionary<Int32, XLCell> columnsCollection))
                {
                    for (int co = 1; co <= MaxColumnUsed; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out XLCell cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }
        }

        internal void SwapRanges(XLSheetRange sheetRange1, XLSheetRange sheetRange2, XLWorksheet worksheet)
        {
            Int32 rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
            Int32 columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    var sp1 = new XLSheetPoint(sheetRange1.FirstPoint.Row + row, sheetRange1.FirstPoint.Column + column);
                    var sp2 = new XLSheetPoint(sheetRange2.FirstPoint.Row + row, sheetRange2.FirstPoint.Column + column);
                    var cell1 = GetCell(sp1);
                    var cell2 = GetCell(sp2);

                    if (cell1 == null) cell1 = worksheet.Cell(sp1.Row, sp1.Column);
                    if (cell2 == null) cell2 = worksheet.Cell(sp2.Row, sp2.Column);

                    // Force evaluation of FormulaR1C1 and clear FormulaA1 - this will preserve the relative formula
                    if (cell1.HasFormula && !string.IsNullOrWhiteSpace(cell1.FormulaR1C1))
                    {
                        cell1.ClearFormulaA1();
                    }

                    cell1.Address = new XLAddress(cell1.Worksheet, sp2.Row, sp2.Column, false, false);
                    Remove(sp1);
                    Add(sp1, cell2);

                    // Force evaluation of FormulaR1C1 and clear FormulaA1 - this will preserve the relative formula
                    if (cell2.HasFormula && !string.IsNullOrWhiteSpace(cell2.FormulaR1C1))
                    {
                        cell2.ClearFormulaA1();
                    }

                    cell2.Address = new XLAddress(cell2.Worksheet, sp1.Row, sp1.Column, false, false);
                    Remove(sp2);
                    Add(sp2, cell1);
                }
            }
        }

        /// <summary/>
        /// <returns>True if the number was lowered to zero so MaxColumnUsed or MaxRowUsed may require
        /// recomputation.</returns>
        private static bool DecrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            if (!dictionary.TryGetValue(key, out Int32 count)) return false;

            if (count > 1)
            {
                dictionary[key] = count - 1;
                return false;
            }
            else
            {
                dictionary.Remove(key);
                return true;
            }
        }

        private static void IncrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            if (dictionary.TryGetValue(key, out Int32 value))
                dictionary[key] = value + 1;
            else
                dictionary.Add(key, 1);
        }

        private void ChangeColumnNumbers(IDictionary<int, int> columnNumberMappings, Dictionary<int, Dictionary<int, XLCell>> rowsCollection, XLRangeAddress affectedRange)
        {
            foreach (var kvp in rowsCollection)
                ChangeColumnNumbers(columnNumberMappings, kvp.Value, affectedRange);
        }

        private Dictionary<int, XLCell> ChangeColumnNumbers(IDictionary<int, int> columnNumberMappings, Dictionary<int, XLCell> cellsCollection, XLRangeAddress affectedRange)
        {
            var currentColumns = columnNumberMappings
                .Where(kvp => cellsCollection.TryGetValue(kvp.Key, out var c)
                              && c.Address.RowNumber >= affectedRange.FirstAddress.RowNumber && c.Address.RowNumber <= affectedRange.LastAddress.RowNumber
                              && c.Address.ColumnNumber != kvp.Value)
                .Select(kvp =>
                {
                    var c = cellsCollection[kvp.Key];
                    // Force evaluation of FormulaR1C1 and clear FormulaA1 - this will preserve the relative formula
                    if (c.HasFormula && !string.IsNullOrWhiteSpace(c.FormulaR1C1)) c.ClearFormulaA1();
                    c.Address = new XLAddress(c.Worksheet, c.Address.RowNumber, kvp.Value, c.Address.FixedRow, c.Address.FixedColumn);

                    return new
                    {
                        OldColumnNumber = kvp.Key,
                        NewColumnNumber = kvp.Value,
                        Cell = c
                    };
                }).ToArray();

            currentColumns.Select(c => c.OldColumnNumber).ForEach(c => cellsCollection.Remove(c));
            currentColumns.ForEach(c => cellsCollection.Add(c.NewColumnNumber, c.Cell));

            return cellsCollection;
        }

        private Dictionary<int, Dictionary<int, XLCell>> ChangeRowNumbers(IDictionary<int, int> rowNumberMappings, Dictionary<int, Dictionary<int, XLCell>> rowsCollection, XLRangeAddress affectedRange)
        {
            var currentRows = rowNumberMappings
                .Select(kvp =>
                {
                    var cellsCollection = rowsCollection[kvp.Key];

                    // Change the row numbers
                    cellsCollection.Values.ForEach(c =>
                    {
                        if (c.Address.ColumnNumber >= affectedRange.FirstAddress.ColumnNumber && c.Address.ColumnNumber <= affectedRange.LastAddress.ColumnNumber &&
                            rowNumberMappings.TryGetValue(c.Address.RowNumber, out var newRowNumber) && c.Address.RowNumber != newRowNumber)
                        {
                            // Force evaluation of FormulaR1C1 and clear FormulaA1 - this will preserve the relative formula
                            if (c.HasFormula && !string.IsNullOrWhiteSpace(c.FormulaR1C1))
                            {
                                c.ClearFormulaA1();
                            }

                            c.Address = new XLAddress(c.Worksheet, rowNumberMappings[c.Address.RowNumber], c.Address.ColumnNumber, c.Address.FixedRow, c.Address.FixedColumn);
                        }
                    });

                    return new
                    {
                        OldRowNumber = kvp.Key,
                        NewRowNumber = kvp.Value,
                        Cells = cellsCollection
                    };
                }).ToArray();

            currentRows.Select(r => r.OldRowNumber).ForEach(r => rowsCollection.Remove(r));
            currentRows.ForEach(r => rowsCollection.Add(r.NewRowNumber, r.Cells));

            var cellsToMove = rowsCollection
                .SelectMany(kvp => kvp.Value.Select(kvp2 => new { Row = kvp.Key, Column = kvp2.Key, Cell = kvp2.Value }))
                .Where(a => a.Row != a.Cell.Address.RowNumber || a.Column != a.Cell.Address.ColumnNumber)
                .ToArray();

            cellsToMove.ForEach(c =>
            {
                rowsCollection[c.Row].Remove(c.Column);
            });

            cellsToMove.ForEach(c =>
            {
                Add(c.Cell.Address.RowNumber, c.Cell.Address.ColumnNumber, c.Cell);
            });
            return rowsCollection;
        }
    }
}
