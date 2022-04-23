using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {
        internal Dictionary<int, int> ColumnsUsed { get; } = new Dictionary<int, int>();
        internal Dictionary<int, HashSet<int>> Deleted { get; } = new Dictionary<int, HashSet<int>>();
        internal Dictionary<int, Dictionary<int, XLCell>> RowsCollection { get; } = new Dictionary<int, Dictionary<int, XLCell>>();

        public int MaxColumnUsed;
        public int MaxRowUsed;
        public Dictionary<int, int> RowsUsed = new Dictionary<int, int>();

        public XLCellsCollection()
        {
            Clear();
        }

        public int Count { get; private set; }

        public void Add(XLSheetPoint sheetPoint, XLCell cell)
        {
            Add(sheetPoint.Row, sheetPoint.Column, cell);
        }

        public void Add(int row, int column, XLCell cell)
        {
            Count++;

            IncrementUsage(RowsUsed, row);
            IncrementUsage(ColumnsUsed, column);

            if (!RowsCollection.TryGetValue(row, out var columnsCollection))
            {
                columnsCollection = new Dictionary<int, XLCell>();
                RowsCollection.Add(row, columnsCollection);
            }
            columnsCollection.Add(column, cell);
            if (row > MaxRowUsed) MaxRowUsed = row;
            if (column > MaxColumnUsed) MaxColumnUsed = column;

            if (Deleted.TryGetValue(row, out var delHash))
                delHash.Remove(column);
        }

        private static void IncrementUsage(Dictionary<int, int> dictionary, int key)
        {
            if (dictionary.TryGetValue(key, out var value))
                dictionary[key] = value + 1;
            else
                dictionary.Add(key, 1);
        }

        /// <summary/>
        /// <returns>True if the number was lowered to zero so MaxColumnUsed or MaxRowUsed may require
        /// recomputation.</returns>
        private static bool DecrementUsage(Dictionary<int, int> dictionary, int key)
        {
            if (!dictionary.TryGetValue(key, out var count)) return false;

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

        public void Clear()
        {
            Count = 0;
            RowsUsed.Clear();
            ColumnsUsed.Clear();

            RowsCollection.Clear();
            MaxRowUsed = 0;
            MaxColumnUsed = 0;
        }

        public void Remove(XLSheetPoint sheetPoint)
        {
            Remove(sheetPoint.Row, sheetPoint.Column);
        }

        public void Remove(int row, int column)
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

            if (Deleted.TryGetValue(row, out var delHash))
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

            if (RowsCollection.TryGetValue(row, out var columnsCollection))
            {
                columnsCollection.Remove(column);
                if (columnsCollection.Count == 0)
                {
                    RowsCollection.Remove(row);
                }
            }
        }

        internal IEnumerable<XLCell> GetCells(int rowStart, int columnStart,
                                            int rowEnd, int columnEnd,
                                            Func<IXLCell, bool> predicate = null)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }
        }

        public int FirstRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, bool> predicate = null)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            return ro;
                    }
                }
            }

            return 0;
        }

        public int FirstColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, bool> predicate = null)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            var firstColumnUsed = finalColumn;
            var found = false;
            for (var ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = columnStart; co <= firstColumnUsed; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
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

        public int LastRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, bool> predicate = null)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = finalRow; ro >= rowStart; ro--)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = finalColumn; co >= columnStart; co--)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            return ro;
                    }
                }
            }
            return 0;
        }

        public int LastColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, XLCellsUsedOptions options,
            Func<IXLCell, bool> predicate = null)
        {
            var maxCo = 0;
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = finalRow; ro >= rowStart; ro--)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = finalColumn; co >= columnStart && co > maxCo; co--)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
                            && !cell.IsEmpty(options)
                            && (predicate == null || predicate(cell)))

                            maxCo = co;
                    }
                }
            }
            return maxCo;
        }

        public void RemoveAll(int rowStart, int columnStart,
                              int rowEnd, int columnEnd)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.ContainsKey(co))
                            Remove(ro, co);
                    }
                }
            }
        }

        public IEnumerable<XLSheetPoint> GetSheetPoints(int rowStart, int columnStart,
                                                        int rowEnd, int columnEnd)
        {
            var finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            var finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (var ro = rowStart; ro <= finalRow; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = columnStart; co <= finalColumn; co++)
                    {
                        if (columnsCollection.ContainsKey(co))
                            yield return new XLSheetPoint(ro, co);
                    }
                }
            }
        }

        public XLCell GetCell(int row, int column)
        {
            if (row > MaxRowUsed || column > MaxColumnUsed)
                return null;

            if (RowsCollection.TryGetValue(row, out var columnsCollection))
            {
                return columnsCollection.TryGetValue(column, out var cell) ? cell : null;
            }
            return null;
        }

        public XLCell GetCell(XLSheetPoint sp)
        {
            return GetCell(sp.Row, sp.Column);
        }

        internal void SwapRanges(XLSheetRange sheetRange1, XLSheetRange sheetRange2, XLWorksheet worksheet)
        {
            var rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
            var columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    var sp1 = new XLSheetPoint(sheetRange1.FirstPoint.Row + row, sheetRange1.FirstPoint.Column + column);
                    var sp2 = new XLSheetPoint(sheetRange2.FirstPoint.Row + row, sheetRange2.FirstPoint.Column + column);
                    var cell1 = GetCell(sp1);
                    var cell2 = GetCell(sp2);

                    if (cell1 == null) cell1 = worksheet.Cell(sp1.Row, sp1.Column);
                    if (cell2 == null) cell2 = worksheet.Cell(sp2.Row, sp2.Column);

                    //if (cell1 != null)
                    //{
                    cell1.Address = new XLAddress(cell1.Worksheet, sp2.Row, sp2.Column, false, false);
                    Remove(sp1);
                    //if (cell2 != null)
                    Add(sp1, cell2);
                    //}

                    //if (cell2 == null) continue;

                    cell2.Address = new XLAddress(cell2.Worksheet, sp1.Row, sp1.Column, false, false);
                    Remove(sp2);
                    //if (cell1 != null)
                    Add(sp2, cell1);
                }
            }
        }

        internal IEnumerable<XLCell> GetCells()
        {
            return GetCells(1, 1, MaxRowUsed, MaxColumnUsed);
        }

        internal IEnumerable<XLCell> GetCells(Func<IXLCell, bool> predicate)
        {
            for (var ro = 1; ro <= MaxRowUsed; ro++)
            {
                if (RowsCollection.TryGetValue(ro, out var columnsCollection))
                {
                    for (var co = 1; co <= MaxColumnUsed; co++)
                    {
                        if (columnsCollection.TryGetValue(co, out var cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }
        }

        public bool Contains(int row, int column)
        {
            return RowsCollection.TryGetValue(row, out var columnsCollection)
                && columnsCollection.ContainsKey(column);
        }

        public int MinRowInColumn(int column)
        {
            for (var row = 1; row <= MaxRowUsed; row++)
            {
                if (RowsCollection.TryGetValue(row, out var columnsCollection)
                    && columnsCollection.ContainsKey(column))

                    return row;
            }

            return 0;
        }

        public int MaxRowInColumn(int column)
        {
            for (var row = MaxRowUsed; row >= 1; row--)
            {
                if (RowsCollection.TryGetValue(row, out var columnsCollection)
                    && columnsCollection.ContainsKey(column))

                    return row;
            }

            return 0;
        }

        public int MinColumnInRow(int row)
        {
            if (RowsCollection.TryGetValue(row, out var columnsCollection)
                && columnsCollection.Any())

                return columnsCollection.Keys.Min();

            return 0;
        }

        public int MaxColumnInRow(int row)
        {
            if (RowsCollection.TryGetValue(row, out var columnsCollection)
                && columnsCollection.Any())

                return columnsCollection.Keys.Max();

            return 0;
        }

        public IEnumerable<XLCell> GetCellsInColumn(int column)
        {
            return GetCells(1, column, MaxRowUsed, column);
        }

        public IEnumerable<XLCell> GetCellsInRow(int row)
        {
            return GetCells(row, 1, row, MaxColumnUsed);
        }
    }
}