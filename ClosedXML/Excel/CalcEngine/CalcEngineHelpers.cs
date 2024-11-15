namespace ClosedXML.Excel.CalcEngine
{
    internal static class CalcEngineHelpers
    {
        internal static bool ValueIsBlank(object? value)
        {
            if (value is null)
                return true;

            if (value is string s)
                return s.Length == 0;

            return false;
        }

        /// <summary>
        /// Get total count of cells in the specified range without initializing them all
        /// (which might cause serious performance issues on column-wide calculations).
        /// </summary>
        /// <param name="rangeExpression">Expression referring to the cell range.</param>
        /// <returns>Total number of cells in the range.</returns>
        internal static long GetTotalCellsCount(XObjectExpression? rangeExpression)
        {
            var (columnCount, rowCount) = GetRangeDimensions(rangeExpression);
            return (long)columnCount * (long)rowCount;
        }

        /// <summary>
        /// Get dimensions of the specified range without initializing them all
        /// (which might cause serious performance issues on column-wide calculations).
        /// </summary>
        /// <param name="rangeExpression">Expression referring to the cell range.</param>
        /// <returns>A tuple of column and row counts.</returns>
        private static (int ColumnCount, int RowCount) GetRangeDimensions(XObjectExpression? rangeExpression)
        {
            var range = (rangeExpression?.Value as CellRangeReference)?.Range;
            if (range == null)
            {
                return (0, 0);
            }

            return (range.ColumnCount(), range.RowCount());
        }

        internal static bool TryExtractRange(Expression expression, out IXLRange? range, out XLError calculationErrorType)
        {
            range = null;
            calculationErrorType = default;

            if (expression is not XObjectExpression objectExpression)
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            if (objectExpression.Value is not CellRangeReference cellRangeReference)
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            range = cellRangeReference.Range;
            return true;
        }
    }
}
