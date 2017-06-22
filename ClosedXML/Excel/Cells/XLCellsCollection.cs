using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {

        private readonly Dictionary<int, Dictionary<int, XLCell>> rowsCollection = new Dictionary<int, Dictionary<int, XLCell>>();
        public readonly Dictionary<Int32, Int32> ColumnsUsed = new Dictionary<int, int>();
        public readonly Dictionary<Int32, HashSet<Int32>> deleted = new Dictionary<int, HashSet<int>>();

        internal Dictionary<int, Dictionary<int, XLCell>> RowsCollection
        {
            get { return rowsCollection; }
        }

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

            Dictionary<int, XLCell> columnsCollection;
            if (!rowsCollection.TryGetValue(row, out columnsCollection))
            {
                columnsCollection = new Dictionary<int, XLCell>();
                rowsCollection.Add(row, columnsCollection);
            }
            columnsCollection.Add(column, cell);
            if (row > MaxRowUsed) MaxRowUsed = row;
            if (column > MaxColumnUsed) MaxColumnUsed = column;

            HashSet<Int32> delHash;
            if (deleted.TryGetValue(row, out delHash))
                delHash.Remove(column);
        }

        private static void IncrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            if (dictionary.ContainsKey(key))
                dictionary[key]++;
            else
                dictionary.Add(key, 1);
        }

        private static void DecrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            Int32 count;
            if (!dictionary.TryGetValue(key, out count)) return;

            if (count > 0)
                dictionary[key]--;
            else
                dictionary.Remove(key);
        }

        public void Clear()
        {
            Count = 0;
            RowsUsed.Clear();
            ColumnsUsed.Clear();

            rowsCollection.Clear();
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
            DecrementUsage(RowsUsed, row);
            DecrementUsage(ColumnsUsed, row);
            
            HashSet<Int32> delHash;
            if (deleted.TryGetValue(row, out delHash))
            {
                if (!delHash.Contains(column))
                    delHash.Add(column);
            }
            else
            {
                delHash = new HashSet<int>();
                delHash.Add(column);
                deleted.Add(row, delHash);
            }

            Dictionary<int, XLCell> columnsCollection;
            if (rowsCollection.TryGetValue(row, out columnsCollection))
            {
                columnsCollection.Remove(column);
                if (columnsCollection.Count == 0)
                {
                    rowsCollection.Remove(row);
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
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }
        }

        internal HashSet<Int32> GetStyleIds(Int32 initial)
        {
            HashSet<Int32> ids = new HashSet<int>();
            ids.Add(initial);
            foreach (var row in rowsCollection)
            {
                foreach (var column in row.Value)
                {
                    var id = column.Value.GetStyleId();
                    if (!ids.Contains(id))
                    {
                        ids.Add(id);
                    }
                }
            }
            return ids;
        }


        internal IEnumerable<XLCell> GetCellsUsed(Int32 rowStart, Int32 columnStart,
                                    Int32 rowEnd, Int32 columnEnd,
                                    Boolean includeFormats, Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                        && !cell.IsEmpty(includeFormats)
                        && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }

        }

        public XLSheetPoint FirstPointUsed(Int32 rowStart, Int32 columnStart,
                                    Int32 rowEnd, Int32 columnEnd, Boolean includeFormats = false, Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;

            var firstRow = FirstRowUsed(rowStart, columnStart, finalRow, finalColumn, includeFormats, predicate);
            if (firstRow == 0) return new XLSheetPoint(0, 0);

            var firstColumn = FirstColumnUsed(rowStart, columnStart, finalRow, finalColumn, includeFormats, predicate);
            if (firstColumn == 0) return new XLSheetPoint(0, 0);

            return new XLSheetPoint(firstRow, firstColumn);
        }

        public XLSheetPoint LastPointUsed(Int32 rowStart, Int32 columnStart,
                            Int32 rowEnd, Int32 columnEnd, Boolean includeFormats = false, Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;

            var firstRow = LastRowUsed(rowStart, columnStart, finalRow, finalColumn, includeFormats, predicate);
            if (firstRow == 0) return new XLSheetPoint(0, 0);

            var firstColumn = LastColumnUsed(rowStart, columnStart, finalRow, finalColumn, includeFormats, predicate);
            if (firstColumn == 0) return new XLSheetPoint(0, 0);

            return new XLSheetPoint(firstRow, firstColumn);
        }

        public int FirstRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, Boolean includeFormats,
            Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                        && !cell.IsEmpty(includeFormats)
                        && (predicate == null || predicate(cell)))
                            return ro;
                    }
                }
            }

            return 0;
        }

        public int FirstColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, Boolean includeFormats, Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = columnStart; co <= finalColumn; co++)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                        && !cell.IsEmpty(includeFormats)
                        && (predicate == null || predicate(cell)))
                            return co;
                    }
                }
            }

            return 0;
        }


        public int LastRowUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, Boolean includeFormats, Func<IXLCell, Boolean> predicate = null)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = finalRow; ro >= rowStart; ro--)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = finalColumn; co >= columnStart; co--)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                         && !cell.IsEmpty(includeFormats)
                        && (predicate == null || predicate(cell)))
                            return ro;
                    }
                }
            }
            return 0;
        }

        public int LastColumnUsed(int rowStart, int columnStart, int rowEnd, int columnEnd, Boolean includeFormats, Func<IXLCell, Boolean> predicate = null)
        {
            int maxCo = 0;
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = finalRow; ro >= rowStart; ro--)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = finalColumn; co >= columnStart && co > maxCo; co--)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                         && !cell.IsEmpty(includeFormats)
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
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
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
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
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
            Dictionary<int, XLCell> columnsCollection;
            if (rowsCollection.TryGetValue(row, out columnsCollection))
            {
                XLCell cell;
                return columnsCollection.TryGetValue(column, out cell) ? cell : null;
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
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(ro, out columnsCollection))
                {
                    for (int co = 1; co <= MaxColumnUsed; co++)
                    {
                        XLCell cell;
                        if (columnsCollection.TryGetValue(co, out cell)
                            && (predicate == null || predicate(cell)))
                            yield return cell;
                    }
                }
            }

        }

        public Boolean Contains(Int32 row, Int32 column)
        {
            Dictionary<int, XLCell> columnsCollection;
            return rowsCollection.TryGetValue(row, out columnsCollection) && columnsCollection.ContainsKey(column);
        }

        public Int32 MinRowInColumn(Int32 column)
        {
            for (int row = 1; row <= MaxRowUsed; row++)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(row, out columnsCollection))
                {
                    if (columnsCollection.ContainsKey(column))
                        return row;
                }
            }

            return 0;
        }

        public Int32 MaxRowInColumn(Int32 column)
        {
            for (int row = MaxRowUsed; row >= 1; row--)
            {
                Dictionary<int, XLCell> columnsCollection;
                if (rowsCollection.TryGetValue(row, out columnsCollection))
                {
                    if (columnsCollection.ContainsKey(column))
                        return row;
                }
            }

            return 0;
        }

        public Int32 MinColumnInRow(Int32 row)
        {
            Dictionary<int, XLCell> columnsCollection;
            if (rowsCollection.TryGetValue(row, out columnsCollection))
            {
                if (columnsCollection.Count > 0)
                    return columnsCollection.Keys.Min();
            }

            return 0;
        }

        public Int32 MaxColumnInRow(Int32 row)
        {
            Dictionary<int, XLCell> columnsCollection;
            if (rowsCollection.TryGetValue(row, out columnsCollection))
            {
                if (columnsCollection.Count > 0)
                    return columnsCollection.Keys.Max();
            }

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