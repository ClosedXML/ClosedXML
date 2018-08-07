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
        #region Public Constructors

        public XLRangeConsolidationEngine(IXLRanges ranges)
        {
            if (ranges == null)
                throw new ArgumentNullException(nameof(ranges));
            _allRanges = ranges;
        }

        #endregion Public Constructors

        #region Public Methods

        public IXLRanges Consolidate()
        {
            if (!_allRanges.Any())
                return _allRanges;

            var worksheets = _allRanges.Select(r => r.Worksheet).Distinct().OrderBy(ws => ws.Position);

            IXLRanges retVal = new XLRanges();
            foreach (var ws in worksheets)
            {
                var matrix = new XLRangeConsolidationMatrix(ws, _allRanges.Where(r => r.Worksheet == ws));
                var consRanges = matrix.GetConsolidatedRanges();
                foreach (var consRange in consRanges)
                {
                    retVal.Add(consRange);
                }
            }

            return retVal;
        }

        #endregion Public Methods

        #region Private Fields

        private readonly IXLRanges _allRanges;

        #endregion Private Fields

        #region Private Classes

        /// <summary>
        /// Class representing the area covering ranges to be consolidated as a set of bit matrices. Does all the dirty job
        /// of ranges consolidation.
        /// </summary>
        private class XLRangeConsolidationMatrix
        {
            #region Public Constructors

            /// <summary>
            /// Constructor.
            /// </summary>
            /// <param name="worksheet">Current worksheet.</param>
            /// <param name="ranges">Ranges to be consolidated. They are expected to belong to the current worksheet, no check is performed.</param>
            public XLRangeConsolidationMatrix(IXLWorksheet worksheet, IEnumerable<IXLRange> ranges)
            {
                _worksheet = worksheet;
                PrepareBitMatrix(ranges);
                FillBitMatrix(ranges);
            }

            #endregion Public Constructors

            #region Public Methods

            /// <summary>
            /// Get consolidated ranges equivalent to the input ones.
            /// </summary>
            public IEnumerable<IXLRange> GetConsolidatedRanges()
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

                        yield return _worksheet.Range(startRow, startColumn, endRow, endColumn);

                        while (j > i)
                        {
                            ClearRangeInRow(_bitMatrix[rowNumbers[j - 1]], starting);
                            j--;
                        }
                    }
                }
            }

            #endregion Public Methods

            #region Private Fields

            private readonly IXLWorksheet _worksheet;
            private Dictionary<int, BitArray> _bitMatrix;
            private int _maxColumn = 0;
            private int _minColumn = XLHelper.MaxColumnNumber + 1;

            #endregion Private Fields

            #region Private Methods

            private void AddToBitMatrix(IXLRangeAddress rangeAddress)
            {
                var rows = _bitMatrix.Keys
                    .Where(k => k >= rangeAddress.FirstAddress.RowNumber &&
                                k <= rangeAddress.LastAddress.RowNumber);

                var minIndex = rangeAddress.FirstAddress.ColumnNumber - _minColumn + 1;
                var maxIndex = rangeAddress.LastAddress.ColumnNumber - _minColumn + 1;

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

            private void FillBitMatrix(IEnumerable<IXLRange> ranges)
            {
                foreach (var range in ranges)
                {
                    AddToBitMatrix(range.RangeAddress);
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

            private void PrepareBitMatrix(IEnumerable<IXLRange> ranges)
            {
                _bitMatrix = new Dictionary<int, BitArray>();
                foreach (var range in ranges)
                {
                    var address = range.RangeAddress;
                    _minColumn = (_minColumn <= address.FirstAddress.ColumnNumber)
                        ? _minColumn
                        : address.FirstAddress.ColumnNumber;
                    _maxColumn = (_maxColumn >= address.LastAddress.ColumnNumber)
                        ? _maxColumn
                        : address.LastAddress.ColumnNumber;

                    if (!_bitMatrix.ContainsKey(address.FirstAddress.RowNumber))
                        _bitMatrix.Add(address.FirstAddress.RowNumber, null);
                    if (!_bitMatrix.ContainsKey(address.LastAddress.RowNumber))
                        _bitMatrix.Add(address.LastAddress.RowNumber, null);
                    if (!_bitMatrix.ContainsKey(address.LastAddress.RowNumber + 1))
                        _bitMatrix.Add(address.LastAddress.RowNumber + 1, null);
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

            #endregion Private Methods
        }

        #endregion Private Classes
    }
}
