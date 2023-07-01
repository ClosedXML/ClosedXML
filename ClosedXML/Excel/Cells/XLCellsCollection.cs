using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {
        private readonly XLWorksheet _ws;
        private readonly List<ISlice> _slices;

        public XLCellsCollection(XLWorksheet ws)
        {
            _ws = ws;
            ValueSlice = new ValueSlice(ws.Workbook.SharedStringTable);
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

        internal Slice<XLCellFormula> FormulaSlice { get; } = new();

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
            var enumerator = new CellsEnumerator(range, this);

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

        internal void SwapRanges(XLSheetRange sheetRange1, XLSheetRange sheetRange2)
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
                var enumerator = new CellsEnumerator(new XLSheetRange(range.FirstPoint.Row, columnNumber, range.LastPoint.Row, columnNumber), this);
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
            var enumerator = new CellsEnumerator(searchRange, this, reverse);

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

        private struct CellsEnumerator
        {
            private readonly List<IEnumerator<XLSheetPoint>> _enumerators;
            private readonly bool _reverse;

            public CellsEnumerator(XLSheetRange range, XLCellsCollection cellsCollection, bool reverse = false)
            {
                Current = new XLSheetPoint(1, 1);
                _reverse = reverse;
                var valueEnumerator = cellsCollection.ValueSlice.GetEnumerator(range, reverse);
                var formulaEnumerator = cellsCollection.FormulaSlice.GetEnumerator(range, reverse);
                var styleEnumerator = cellsCollection.StyleSlice.GetEnumerator(range, reverse);
                var kitchenSinkEnumerator = cellsCollection.MiscSlice.GetEnumerator(range, reverse);

                _enumerators = new();
                if (valueEnumerator.MoveNext())
                    _enumerators.Add(valueEnumerator);
                if (formulaEnumerator.MoveNext())
                    _enumerators.Add(formulaEnumerator);
                if (styleEnumerator.MoveNext())
                    _enumerators.Add(styleEnumerator);
                if (kitchenSinkEnumerator.MoveNext())
                    _enumerators.Add(kitchenSinkEnumerator);
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
    }
}
