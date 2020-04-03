using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {
        internal Dictionary<Int32, Int32> ColumnsUsed { get; } = new Dictionary<int, int>();
        internal Dictionary<Int32, HashSet<Int32>> Deleted { get; } = new Dictionary<int, HashSet<int>>();
        internal Dictionary<int, Dictionary<int, XLCell>> RowsCollection { get; } = new Dictionary<int, Dictionary<int, XLCell>>();

        public Int32 MaxColumnUsed;
        public Int32 MaxRowUsed;
        public Dictionary<Int32, Int32> RowsUsed = new Dictionary<int, int>();

        public XLCellsCollection()
        {
            Clear();
        }

        public Int32 Count { get; private set; }

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

        private static void IncrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            if (dictionary.TryGetValue(key, out Int32 value))
                dictionary[key] = value + 1;
            else
                dictionary.Add(key, 1);
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

        public Boolean Contains(Int32 row, Int32 column)
        {
            return RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                && columnsCollection.ContainsKey(column);
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

        public Int32 MaxColumnInRow(Int32 row)
        {
            if (RowsCollection.TryGetValue(row, out Dictionary<Int32, XLCell> columnsCollection)
                && columnsCollection.Any())

                return columnsCollection.Keys.Max();

            return 0;
        }

        public IEnumerable<XLCell> GetCellsInColumn(Int32 column)
        {
            return GetCells(1, column, MaxRowUsed, column);
        }

        public IEnumerable<XLCell> GetCellsInRow(Int32 row)
        {
            return GetCells(row, 1, row, MaxColumnUsed);
        }
    }
}
