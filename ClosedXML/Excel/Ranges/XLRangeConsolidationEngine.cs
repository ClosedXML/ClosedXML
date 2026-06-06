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
                var areaList = new XLAreaList(_allRanges.Where<XLRange>(r => r.Worksheet == ws).Select(r => r.SheetRange).ToList());
                var matrix = new XLRangeConsolidationMatrix(areaList);
                var consRanges = matrix.GetConsolidatedRanges();
                foreach (var consArea in consRanges)
                {
                    retVal.Add(ws.Range(consArea));
                }
            }

            return retVal;
        }

        internal static XLAreaList Consolidate(XLAreaList areas)
        {
            if (areas.Count == 0)
                return areas;

            var matrix = new XLRangeConsolidationMatrix(areas);
            var consRanges = matrix.GetConsolidatedRanges().ToList();
            return new XLAreaList(consRanges);
        }

        /// <summary>
        /// Class representing the area covering ranges to be consolidated as a set of bit matrices. Does all the dirty job
        /// of ranges consolidation.
        /// </summary>
        private class XLRangeConsolidationMatrix
        {
            private readonly Dictionary<int, BitArray> _bitMatrix;
            private readonly int _minColumn;

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="areas">Areas to be consolidated.</param>
            internal XLRangeConsolidationMatrix(XLAreaList areas)
            {
                (_bitMatrix, _minColumn) = PrepareBitMatrix(areas);
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
                    .Where(k => k >= area.TopRow &&
                                k <= area.BottomRow);

                var minIndex = area.LeftColumn - _minColumn + 1;
                var maxIndex = area.RightColumn - _minColumn + 1;

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

            private static (Dictionary<int, BitArray> BitMatrix, int MinColumn) PrepareBitMatrix(XLAreaList areas)
            {
                var minColumn = XLHelper.MaxColumnNumber + 1;
                var maxColumn = 0;
                foreach (var area in areas)
                {
                    minColumn = (minColumn <= area.LeftColumn)
                        ? minColumn
                        : area.LeftColumn;
                    maxColumn = (maxColumn >= area.RightColumn)
                        ? maxColumn
                        : area.RightColumn;
                }

                var bitMaskSize = maxColumn - minColumn + 3;
                var bitMatrix = new Dictionary<int, BitArray>();
                foreach (var area in areas)
                {
                    AddRowBitmask(bitMatrix, area.TopRow, bitMaskSize);
                    AddRowBitmask(bitMatrix, area.BottomRow, bitMaskSize);
                    AddRowBitmask(bitMatrix, area.BottomRow + 1, bitMaskSize);
                }

                return (bitMatrix, minColumn);

                static void AddRowBitmask(Dictionary<int, BitArray> bitMatrix, int rowNum, int bitMaskSize)
                {
                    if (!bitMatrix.ContainsKey(rowNum))
                        bitMatrix.Add(rowNum, new BitArray(bitMaskSize, false));
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
