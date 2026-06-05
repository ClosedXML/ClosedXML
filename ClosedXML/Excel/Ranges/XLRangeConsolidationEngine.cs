#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Engine for ranges consolidation. Supports IXLRanges including ranges from either one or multiple worksheets.
    /// </summary>
    internal class XLRangeConsolidationEngine
    {
        private readonly XLWorkbook _workbook;
        private readonly XLRanges _allRanges;

        public XLRangeConsolidationEngine(XLWorkbook workbook, XLRanges ranges)
        {
            _workbook = workbook;
            _allRanges = ranges ?? throw new ArgumentNullException(nameof(ranges));
        }

        public XLRanges Consolidate()
        {
            if (_allRanges.Count == 0)
                return _allRanges;

            var worksheets = _allRanges.Select<XLRange, XLWorksheet>(r => r.Worksheet).Distinct().OrderBy(ws => ws.Position);

            var retVal = new XLRanges(_workbook);
            foreach (var ws in worksheets)
            {
                var matrix = new XLRangeConsolidationMatrix(_allRanges.Where<XLRange>(r => r.Worksheet == ws).Select(r => r.SheetRange).ToList());
                var consRanges = matrix.GetConsolidatedRanges();
                foreach (var consArea in consRanges)
                {
                    retVal.Add(ws.Range(consArea));
                }
            }

            return retVal;
        }

        /// <summary>
        /// Class representing the area covering ranges to be consolidated as a set of bit matrices. Does all the dirty job
        /// of ranges consolidation.
        /// </summary>
        private class XLRangeConsolidationMatrix
        {
            private Dictionary<int, BitArray> _bitMatrix;
            private int _maxColumn = 0;
            private int _minColumn = XLHelper.MaxColumnNumber + 1;

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="areas">Areas to be consolidated.</param>
            internal XLRangeConsolidationMatrix(IReadOnlyCollection<XLSheetRange> areas)
            {
                PrepareBitMatrix(areas);
                FillBitMatrix(areas);
            }

            /// <summary>
            /// Get consolidated ranges equivalent to the input ones.
            /// </summary>
            public IEnumerable<XLSheetRange> GetConsolidatedRanges()
            {
                var rowNumbers = _bitMatrix.Keys.OrderBy(k => k).ToArray();
                for (int i = 0; i < rowNumbers.Length; i++)
                {
                    var startRow = rowNumbers[i];
                    var startings = GetRangesBoundariesStartingByRow(_bitMatrix[startRow]);

                    foreach (var starting in startings)
                    {
                        int j = i + 1;
                        while (j < rowNumbers.Length && RowIncludesRange(_bitMatrix[rowNumbers[j]], starting)) j++;

                        var endRow = rowNumbers[j - 1];
                        var startColumn = starting.Item1 + _minColumn - 1;
                        var endColumn = starting.Item2 + _minColumn - 1;

                        yield return new XLSheetRange(startRow, startColumn, endRow, endColumn);

                        while (j > i)
                        {
                            ClearRangeInRow(_bitMatrix[rowNumbers[j - 1]], starting);
                            j--;
                        }
                    }
                }
            }

            private void AddToBitMatrix(XLSheetRange area)
            {
                var rows = _bitMatrix.Keys
                    .Where(k => k >= area.FirstPoint.Row &&
                                k <= area.LastPoint.Row);

                var minIndex = area.FirstPoint.Column - _minColumn + 1;
                var maxIndex = area.LastPoint.Column - _minColumn + 1;

                foreach (var rowNum in rows)
                {
                    for (int i = minIndex; i <= maxIndex; i++)
                    {
                        _bitMatrix[rowNum][i] = true;
                    }
                }
            }

            private void ClearRangeInRow(BitArray rowArray, Tuple<int, int> rangeBoundaries)
            {
                for (int i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
                {
                    rowArray[i] = false;
                }
            }

            private void FillBitMatrix(IEnumerable<XLSheetRange> areas)
            {
                foreach (var area in areas)
                {
                    AddToBitMatrix(area);
                }

                System.Diagnostics.Debug.Assert(
                    _bitMatrix.Values.All(r => r[0] == false && r[r.Length - 1] == false));
            }

            private IEnumerable<Tuple<int, int>> GetRangesBoundariesStartingByRow(BitArray rowArray)
            {
                int startIdx = 0;
                for (int i = 1; i < rowArray.Length - 1; i++)
                {
                    if (!rowArray[i - 1] && rowArray[i])
                        startIdx = i;
                    if (rowArray[i] && !rowArray[i + 1])
                        yield return new Tuple<int, int>(startIdx, i);
                }
            }

            private void PrepareBitMatrix(IEnumerable<XLSheetRange> areas)
            {
                _bitMatrix = new Dictionary<int, BitArray>();
                foreach (var area in areas)
                {
                    _minColumn = (_minColumn <= area.FirstPoint.Column)
                        ? _minColumn
                        : area.FirstPoint.Column;
                    _maxColumn = (_maxColumn >= area.LastPoint.Column)
                        ? _maxColumn
                        : area.LastPoint.Column;

                    if (!_bitMatrix.ContainsKey(area.FirstPoint.Row))
                        _bitMatrix.Add(area.FirstPoint.Row, null);
                    if (!_bitMatrix.ContainsKey(area.LastPoint.Row))
                        _bitMatrix.Add(area.LastPoint.Row, null);
                    if (!_bitMatrix.ContainsKey(area.LastPoint.Row + 1))
                        _bitMatrix.Add(area.LastPoint.Row + 1, null);
                }

                var keys = _bitMatrix.Keys.ToList();
                foreach (var rowNum in keys)
                {
                    _bitMatrix[rowNum] = new BitArray(_maxColumn - _minColumn + 3, false);
                }
            }

            private bool RowIncludesRange(BitArray rowArray, Tuple<int, int> rangeBoundaries)
            {
                for (int i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
                {
                    if (!rowArray[i])
                        return false;
                }

                return true;
            }
        }
    }
}
