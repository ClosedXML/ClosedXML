using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection : IWorkbookListener
    {
        private readonly XLWorksheet _ws;
        private readonly List<ISlice> _slices;

        public XLCellsCollection(XLWorksheet ws)
        {
            _ws = ws;
            ValueSlice = new ValueSlice(ws.Workbook.SharedStringTable);
            FormulaSlice = new FormulaSlice(ws);
            _slices = new List<ISlice> { ValueSlice, FormulaSlice, StyleSlice, MiscSlice };
        }

        internal HashSet<int> ColumnsUsedKeys
        {
            get
            {
                var set = new HashSet<int>();
                foreach (var slice in _slices)
                    set.UnionWith(slice.UsedColumns);

                return set;
            }
        }

        internal bool IsEmpty => _slices.All(slice => slice.IsEmpty);

        internal Int32 MaxColumnUsed
        {
            get
            {
                var max = int.MinValue;
                foreach (var slice in _slices)
                    max = Math.Max(max, slice.MaxColumn);

                return Math.Max(1, max);
            }
        }

        internal Int32 MaxRowUsed
        {
            get
            {
                var max = int.MinValue;
                foreach (var slice in _slices)
                    max = Math.Max(max, slice.MaxRow);

                return Math.Max(1, max);
            }
        }

        internal HashSet<int> RowsUsedKeys
        {
            get
            {
                var set = new HashSet<int>();
                foreach (var slice in _slices)
                    set.UnionWith(slice.UsedRows);

                return set;
            }
        }

        internal ValueSlice ValueSlice { get; }

        internal FormulaSlice FormulaSlice { get; }

        internal Slice<XLStyleValue?> StyleSlice { get; } = new();

        internal Slice<XLMiscSliceContent> MiscSlice { get; } = new();

        internal XLWorksheet Worksheet => _ws;

        internal void Clear()
        {
            Clear(XLSheetRange.Full);
        }

        internal void Clear(XLSheetRange clearRange)
        {
            foreach (var slice in _slices)
                slice.Clear(clearRange);
        }

        internal void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete)
        {
            foreach (var slice in _slices)
                slice.DeleteAreaAndShiftLeft(rangeToDelete);
        }

        internal void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete)
        {
            foreach (var slice in _slices)
                slice.DeleteAreaAndShiftUp(rangeToDelete);
        }

        internal XLCell GetCell(XLSheetPoint address)
        {
            return new XLCell(_ws, address);
        }

        /// <summary>
        /// Get all used cells in the worksheet.
        /// </summary>
        internal IEnumerable<XLCell> GetCells()
        {
            return GetCells(XLSheetRange.Full);
        }

        /// <summary>
        /// Get all used cells in the worksheet that satisfy the predicate.
        /// </summary>
        internal IEnumerable<XLCell> GetCells(Func<XLCell, Boolean> predicate)
        {
            return GetCells(XLSheetRange.Full, predicate);
        }

        /// <summary>
        /// Get all used cells in the range that satisfy the predicate.
        /// </summary>
        internal IEnumerable<XLCell> GetCells(Int32 rowStart, Int32 columnStart,
                                            Int32 rowEnd, Int32 columnEnd,
                                            Func<XLCell, Boolean>? predicate = null)
        {
            return GetCells(new XLSheetRange(rowStart, columnStart, rowEnd, columnEnd), predicate);
        }

        /// <summary>
        /// Get all used cells in the range that satisfy the predicate.
        /// </summary>
        internal IEnumerable<XLCell> GetCells(XLSheetRange range, Func<XLCell, Boolean>? predicate = null)
        {
            var enumerator = new SlicesEnumerator(range, this);

            while (enumerator.MoveNext())
            {
                var cellAddress = enumerator.Current;
                var cell = GetCell(cellAddress);
                if (predicate == null || predicate(cell))
                    yield return cell;
            }
        }

        internal IEnumerable<XLCell> GetCellsInColumn(Int32 column)
        {
            return GetCells(1, column, XLHelper.MaxRowNumber, column);
        }

        internal IEnumerable<XLCell> GetCellsInRow(Int32 row)
        {
            return GetCells(row, 1, row, XLHelper.MaxColumnNumber);
        }

        /// <summary>
        /// Get cell or null, if cell is not used.
        /// </summary>
        internal XLCell? GetUsedCell(XLSheetPoint address)
        {
            if (!IsUsed(address))
                return null;

            return GetCell(address);
        }

        internal int FirstColumnUsed(XLSheetRange searchRange, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate = null)
        {
            return FindUsedColumn(searchRange, options, predicate, false);
        }

        internal int FirstRowUsed(XLSheetRange searchRange, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate = null)
        {
            return FindUsedRow(searchRange, options, predicate, false);
        }

        internal void InsertAreaAndShiftDown(XLSheetRange insertedRange)
        {
            foreach (var slice in _slices)
                slice.InsertAreaAndShiftDown(insertedRange);
        }

        internal void InsertAreaAndShiftRight(XLSheetRange insertedRange)
        {
            foreach (var slice in _slices)
                slice.InsertAreaAndShiftRight(insertedRange);
        }

        internal int LastColumnUsed(XLSheetRange searchRange, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate = null)
        {
            return FindUsedColumn(searchRange, options, predicate, true);
        }

        internal int LastRowUsed(XLSheetRange searchRange, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate = null)
        {
            return FindUsedRow(searchRange, options, predicate, true);
        }

        /// <summary>
        /// Remap rows of a range.
        /// </summary>
        /// <param name="map">A sorted map of rows. The values must be resorted row numbers from <paramref name="sheetRange"/>.</param>
        /// <param name="sheetRange">Sheet that should have its rows rearranged.</param>
        internal void RemapRows(IList<int> map, XLSheetRange sheetRange)
        {
            RemapRanges(map, sheetRange.TopRow, SwapRows);

            void SwapRows(int prevRowNumber, int currentRowNumber)
            {
                var prevRowRange = new XLSheetRange(
                    new XLSheetPoint(prevRowNumber, sheetRange.LeftColumn),
                    new XLSheetPoint(prevRowNumber, sheetRange.RightColumn));
                var currentRowRange = new XLSheetRange(
                    new XLSheetPoint(currentRowNumber, sheetRange.LeftColumn),
                    new XLSheetPoint(currentRowNumber, sheetRange.RightColumn));
                SwapRanges(prevRowRange, currentRowRange);
            }
        }

        /// <summary>
        /// Remap columns of a range.
        /// </summary>
        /// <param name="map">A sorted map of columns. The values must be resorted columns numbers from <paramref name="sheetRange"/>.</param>
        /// <param name="sheetRange">Sheet that should have its columns rearranged.</param>
        internal void RemapColumns(IList<int> map, XLSheetRange sheetRange)
        {
            RemapRanges(map, sheetRange.LeftColumn, SwapColumns);

            void SwapColumns(int prevColNumber, int currentColNumber)
            {
                var prevRowRange = new XLSheetRange(
                    new XLSheetPoint(sheetRange.TopRow, prevColNumber),
                    new XLSheetPoint(sheetRange.BottomRow, prevColNumber));
                var currentRowRange = new XLSheetRange(
                    new XLSheetPoint(sheetRange.TopRow, currentColNumber),
                    new XLSheetPoint(sheetRange.BottomRow, currentColNumber));
                SwapRanges(prevRowRange, currentRowRange);
            }
        }

        private static void RemapRanges(IList<int> map, int indexOffset, Action<int, int> swapData)
        {
            for (var i = 0; i < map.Count; ++i)
            {
                var axisNumber = i + indexOffset;
                var dataAxisNumber = map[i];
                if (axisNumber == dataAxisNumber)
                    continue;

                // Current row doesn't contain data it should, so it is a part of a permutation
                // loop. Go over each item in a loop and 
                // We need to replace
                var prevNumber = axisNumber;
                var currentNumber = dataAxisNumber;
                var startLoopNumber = prevNumber;
                do
                {
                    // Current row number contains data that should be on the previous row number,
                    // so swap them. That will fix another link in a loop (the previous one), but
                    // will keep current inconsistent, but that will be fixed when loop completes.
                    swapData(prevNumber, currentNumber);

                    // Because previous row number is already fixed and will no longer be touched
                    // during loop fix, mark it as a row that contains correct data.
                    map[prevNumber - indexOffset] = prevNumber;

                    prevNumber = currentNumber;
                    currentNumber = map[currentNumber - indexOffset];
                } while (currentNumber != startLoopNumber);

                // Although we don't have to swap the last one (N count loop needs only N-1 swaps),
                // we have to mark the last row mapping for the last link (the one before start).
                map[prevNumber - indexOffset] = prevNumber;
            }
        }

        private void SwapRanges(XLSheetRange sheetRange1, XLSheetRange sheetRange2)
        {
            var rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
            var columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    var sp1 = new XLSheetPoint(sheetRange1.FirstPoint.Row + row, sheetRange1.FirstPoint.Column + column);
                    var sp2 = new XLSheetPoint(sheetRange2.FirstPoint.Row + row, sheetRange2.FirstPoint.Column + column);

                    SwapCellsContent(sp1, sp2);
                }
            }
        }

        private int FindUsedColumn(XLSheetRange range, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate, bool descending)
        {
            var usedColumns = Enumerable.Empty<int>();
            foreach (var slice in _slices)
                usedColumns = usedColumns.Concat(slice.UsedColumns);

            usedColumns = usedColumns
                .Where(c => c >= range.FirstPoint.Column && c <= range.LastPoint.Column)
                .Distinct();
            usedColumns = descending
                ? usedColumns.OrderByDescending(x => x)
                : usedColumns.OrderBy(x => x);

            foreach (var columnNumber in usedColumns)
            {
                var enumerator = new SlicesEnumerator(new XLSheetRange(range.FirstPoint.Row, columnNumber, range.LastPoint.Row, columnNumber), this);
                while (enumerator.MoveNext())
                {
                    var cell = new XLCell(_ws, enumerator.Current);
                    if (!cell.IsEmpty(options) &&
                        (predicate == null || predicate(cell)))
                    {
                        return enumerator.Current.Column;
                    }
                }
            }

            return 0;
        }

        private int FindUsedRow(XLSheetRange searchRange, XLCellsUsedOptions options, Func<IXLCell, Boolean>? predicate, bool reverse)
        {
            var enumerator = new SlicesEnumerator(searchRange, this, reverse);

            while (enumerator.MoveNext())
            {
                var cellAddress = enumerator.Current;
                var cell = GetCell(cellAddress);
                if (!cell.IsEmpty(options)
                    && (predicate == null || predicate(cell)))
                    return cellAddress.Row;
            }

            return 0;
        }

        private bool IsUsed(XLSheetPoint address)
        {
            // This is different from XLCellUsedOptions, which uses a business logic (e.g. empty string is considered not-used).
            // Here, we ask whether any slice contains a used elements which might differ from cell used logic.
            foreach (var slice in _slices)
            {
                if (slice.IsUsed(address))
                    return true;
            }

            return false;
        }

        internal void SwapCellsContent(XLSheetPoint sp1, XLSheetPoint sp2)
        {
            ValueSlice.Swap(sp1, sp2);
            FormulaSlice.Swap(sp1, sp2);
            StyleSlice.Swap(sp1, sp2);
            MiscSlice.Swap(sp1, sp2);
        }

        internal SlicesEnumerator ForValuesAndFormulas(XLSheetRange range)
        {
            var valueEnumerator = ValueSlice.GetEnumerator(range);
            var formulaEnumerator = FormulaSlice.GetEnumerator(range);
            return new SlicesEnumerator(false, valueEnumerator, formulaEnumerator);
        }

        /// <summary>
        /// Enumerator that combines several other slice enumerators and enumerates
        /// <see cref="XLSheetPoint"/> in any of them.
        /// </summary>
        internal struct SlicesEnumerator
        {
            private readonly List<IEnumerator<XLSheetPoint>> _enumerators;
            private readonly bool _reverse;

            public SlicesEnumerator(XLSheetRange range, XLCellsCollection cellsCollection, bool reverse = false)
                : this(
                    reverse,
                    cellsCollection.ValueSlice.GetEnumerator(range, reverse),
                    cellsCollection.FormulaSlice.GetEnumerator(range, reverse),
                    cellsCollection.StyleSlice.GetEnumerator(range, reverse),
                    cellsCollection.MiscSlice.GetEnumerator(range, reverse))
            {
            }

            public SlicesEnumerator(bool reverse, params IEnumerator<XLSheetPoint>[] enumerators)
            {
                Current = new XLSheetPoint(1, 1);
                _reverse = reverse;
                _enumerators = new();
                foreach (var enumerator in enumerators)
                {
                    if (enumerator.MoveNext())
                        _enumerators.Add(enumerator);
                }
            }

            public XLSheetPoint Current { get; private set; }

            public bool MoveNext()
            {
                XLSheetPoint? current = null;
                for (var i = 0; i < _enumerators.Count; ++i)
                {
                    var enumerator = _enumerators[i];
                    if (current is null || (
                            _reverse
                                ? enumerator.Current.CompareTo(current.Value) > 0
                                : enumerator.Current.CompareTo(current.Value) < 0
                            ))
                        current = enumerator.Current;
                }

                if (current == null)
                    return false;

                Current = current.Value;

                for (var i = _enumerators.Count - 1; i >= 0; --i)
                {
                    var enumerator = _enumerators[i];
                    if (enumerator.Current == current)
                    {
                        var isDone = !enumerator.MoveNext();
                        if (isDone)
                        {
                            _enumerators.RemoveAt(i);
                        }
                    }
                }

                return true;
            }
        }

        void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        {
            using var enumerator = FormulaSlice.GetForwardEnumerator(XLSheetRange.Full);
            while (enumerator.MoveNext())
            {
                ref readonly XLCellFormula cellFormula = ref enumerator.Current;
                var currentPoint = enumerator.Point;
                if (cellFormula.Type != FormulaType.Normal)
                {
                    // Array or data formula. Only change name once, on master cell.
                    var isMasterCell = cellFormula.Range.FirstPoint == currentPoint;
                    if (!isMasterCell)
                    {
                        continue;
                    }
                }

                cellFormula.RenameSheet(currentPoint, oldSheetName, newSheetName);
            }
        }
    }
}
