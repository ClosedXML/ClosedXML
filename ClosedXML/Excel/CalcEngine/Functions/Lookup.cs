#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Lookup
    {
        public static void Register(FunctionRegistry ce)
        {
            //ce.RegisterFunction("ADDRESS", , Address); // Returns a reference as text to a single cell in a worksheet
            //ce.RegisterFunction("AREAS", , Areas); // Returns the number of areas in a reference
            //ce.RegisterFunction("CHOOSE", , Choose); // Chooses a value from a list of values
            ce.RegisterFunction("COLUMN", 0, 1, Column, FunctionFlags.Range, AllowRange.All); // Returns the column number of a reference
            ce.RegisterFunction("COLUMNS", 1, 1, Adapt(Columns), FunctionFlags.Range, AllowRange.All); // Returns the number of columns in a reference
            //ce.RegisterFunction("FORMULATEXT", , Formulatext); // Returns the formula at the given reference as text
            //ce.RegisterFunction("GETPIVOTDATA", , Getpivotdata); // Returns data stored in a PivotTable report
            ce.RegisterFunction("HLOOKUP", 3, 4, AdaptLastOptional(Hlookup, true), FunctionFlags.Range, AllowRange.Only, 1); // Looks in the top row of an array and returns the value of the indicated cell
            ce.RegisterFunction("HYPERLINK", 1, 2, Adapt(Hyperlink), FunctionFlags.Scalar | FunctionFlags.SideEffect); // Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
            ce.RegisterFunction("INDEX", 2, 4, Index, AllowRange.Only, 0, 1); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            ce.RegisterFunction("MATCH", 2, 3, Match, AllowRange.Only, 1); // Looks up values in a reference or array
            //ce.RegisterFunction("OFFSET", , Offset); // Returns a reference offset from a given reference
            ce.RegisterFunction("ROW", 0, 1, Row, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Returns the row number of a reference
            ce.RegisterFunction("ROWS", 1, 1, Adapt(Rows), FunctionFlags.Range, AllowRange.All); // Returns the number of rows in a reference
            //ce.RegisterFunction("RTD", , Rtd); // Retrieves real-time data from a program that supports COM automation
            ce.RegisterFunction("TRANSPOSE", 1, 1, Adapt(Transpose), FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Returns the transpose of an array
            ce.RegisterFunction("VLOOKUP", 3, 4, AdaptLastOptional(Vlookup, true), FunctionFlags.Range, AllowRange.Only, 1); // Looks in the first column of an array and moves across the row to return the value of a cell
        }

        private static AnyValue Column(CalcContext ctx, Span<AnyValue> p)
        {
            if (p.Length == 0 || p[0].IsBlank)
                return ctx.FormulaAddress.ColumnNumber;

            if (!p[0].TryPickArea(out var area, out var error))
                return error;

            var firstColumn = area.FirstAddress.ColumnNumber;
            var lastColumn = area.LastAddress.ColumnNumber;
            if (firstColumn == lastColumn)
                return firstColumn;

            var span = lastColumn - firstColumn + 1;
            var array = new ScalarValue[1, span];
            for (var col = firstColumn; col <= lastColumn; col++)
                array[0, col - firstColumn] = col;

            return new ConstArray(array);
        }

        private static AnyValue Columns(CalcContext _, AnyValue value)
        {
            return RowsOrColumns(value, false);
        }

        private static AnyValue Hlookup(CalcContext ctx, ScalarValue lookupValue, AnyValue rangeValue, double rowNumber, bool approximateSearchFlag)
        {
            if (lookupValue.IsError)
                return lookupValue.ToAnyValue();

            // Only the lookup value is converted to 0, not values in the range
            if (lookupValue.IsBlank)
                lookupValue = 0;

            if (lookupValue.TryPickText(out var lookupText, out _) && lookupText.Length > 255)
                return XLError.IncompatibleValue;

            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var array, out var reference))
            {
                if (reference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                array = new ReferenceArray(reference.Areas.Single(), ctx);
            }

            var rowIndex = (int)Math.Truncate(rowNumber) - 1;
            if (rowIndex < 0)
                return XLError.IncompatibleValue;
            if (rowIndex >= array.Height)
                return XLError.CellReference;

            if (approximateSearchFlag)
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var transposedArray = new TransposedArray(array);
                var foundColumn = Bisection(transposedArray, lookupValue);
                if (foundColumn == -1)
                    return XLError.NoValueAvailable;

                return array[rowIndex, foundColumn].ToAnyValue();
            }
            else
            {
                // TODO: Implement wildcard search
                for (var columnIndex = 0; columnIndex < array.Width; columnIndex++)
                {
                    var currentValue = array[0, columnIndex];

                    // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                    var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                    if (comparison == 0)
                        return array[rowIndex, columnIndex].ToAnyValue();
                }

                return XLError.NoValueAvailable;
            }
        }

        private static AnyValue Hyperlink(CalcContext ctx, string linkLocation, ScalarValue? friendlyName)
        {
            return friendlyName?.ToAnyValue() ?? linkLocation;
        }

        private static object Index(List<Expression> p)
        {
            // This is one of the few functions that is "overloaded"
            if (!CalcEngineHelpers.TryExtractRange(p[0], out var range, out var error))
                return error;

            if (range.ColumnCount() > 1 && range.RowCount() > 1)
            {
                var row_num = (int)p[1];
                var column_num = (int)p[2];

                if (row_num > range.RowCount())
                    return XLError.CellReference;

                if (column_num > range.ColumnCount())
                    return XLError.CellReference;

                return range.Row(row_num).Cell(column_num).Value;
            }
            else if (p.Count == 2)
            {
                var cellOffset = (int)p[1];
                if (cellOffset > range.RowCount() * range.ColumnCount())
                    return XLError.CellReference;

                return range.Cells().ElementAt(cellOffset - 1).Value;
            }
            else
            {
                int column_num = 1;
                int row_num = 1;

                if (!(p[1] is EmptyValueExpression))
                    row_num = (int)p[1];

                if (!(p[2] is EmptyValueExpression))
                    column_num = (int)p[2];

                var rangeIsRow = range.RowCount() == 1;
                if (rangeIsRow && row_num > 1)
                    return XLError.CellReference;

                if (!rangeIsRow && column_num > 1)
                    return XLError.CellReference;

                if (row_num > range.RowCount())
                    return XLError.CellReference;

                if (column_num > range.ColumnCount())
                    return XLError.CellReference;

                return range.Row(row_num).Cell(column_num).Value;
            }
        }

        private static object Match(List<Expression> p)
        {
            var lookup_value = p[0];

            if (!CalcEngineHelpers.TryExtractRange(p[1], out var range, out var error))
                return error;

            int match_type = 1;
            if (p.Count > 2)
                match_type = Math.Sign((int)p[2]);

            if (range.ColumnCount() != 1 && range.RowCount() != 1)
                return XLError.IncompatibleValue;

            Predicate<int> lookupPredicate = null;
            switch (match_type)
            {
                case 0:
                    lookupPredicate = i => i == 0;
                    break;

                case 1:
                    lookupPredicate = i => i <= 0;
                    break;

                case -1:
                    lookupPredicate = i => i >= 0;
                    break;

                default:
                    return XLError.NoValueAvailable;
            }

            IXLCell foundCell = null;

            if (match_type == 0)
                foundCell = range
                    .CellsUsed(XLCellsUsedOptions.Contents, c => lookupPredicate.Invoke(new Expression(c.Value).CompareTo(lookup_value)))
                    .FirstOrDefault();
            else
            {
                var isFirst = true;
                XLCellValue previousValue = Blank.Value;
                foundCell = range
                    .CellsUsed(XLCellsUsedOptions.Contents)
                    .TakeWhile(c =>
                    {
                        var currentCellExpression = new Expression(c.Value);
                        if (!isFirst)
                        {
                            // When match_type != 0, we have to assume that the order of the items being search is ascending or descending
                            var previousValueExpression = new Expression(previousValue);
                            if (!lookupPredicate.Invoke(previousValueExpression.CompareTo(currentCellExpression)))
                                return false;
                        }

                        isFirst = false;
                        previousValue = c.Value;

                        return lookupPredicate.Invoke(currentCellExpression.CompareTo(lookup_value));
                    })
                    .LastOrDefault();
            }

            if (foundCell == null)
                return XLError.NoValueAvailable;

            var firstCell = range.FirstCell();

            return (foundCell.Address.ColumnNumber - firstCell.Address.ColumnNumber + 1) * (foundCell.Address.RowNumber - firstCell.Address.RowNumber + 1);
        }

        private static AnyValue Row(CalcContext ctx, Span<AnyValue> p)
        {
            if (p.Length == 0 || p[0].IsBlank)
                return ctx.FormulaAddress.RowNumber;

            if (!p[0].TryPickArea(out var area, out var error))
                return error;

            var firstRow = area.FirstAddress.RowNumber;
            var lastRow = area.LastAddress.RowNumber;
            if (firstRow == lastRow)
                return firstRow;

            var span = lastRow - firstRow + 1;
            var array = new ScalarValue[span, 1];
            for (var row = firstRow; row <= lastRow; row++)
                array[row - firstRow, 0] = row;

            return new ConstArray(array);
        }

        private static AnyValue Rows(CalcContext _, AnyValue value)
        {
            return RowsOrColumns(value, true);
        }

        private static AnyValue Transpose(CalcContext ctx, AnyValue value)
        {
            if (value.TryPickSingleOrMultiValue(out var single, out var multi, ctx))
                return single.ToAnyValue();

            return new TransposedArray(multi);
        }

        private static AnyValue Vlookup(CalcContext ctx, ScalarValue lookupValue, AnyValue rangeValue, double columnNumber, bool approximateSearchFlag)
        {
            if (lookupValue.IsError)
                return lookupValue.ToAnyValue();

            // Only the lookup value is converted to 0, not values in the range
            if (lookupValue.IsBlank)
                lookupValue = 0;

            if (lookupValue.TryPickText(out var lookupText, out _) && lookupText.Length > 255)
                return XLError.IncompatibleValue;

            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var array, out var reference))
            {
                if (reference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                array = new ReferenceArray(reference.Areas.Single(), ctx);
            }

            var columnIdx = (int)Math.Truncate(columnNumber) - 1;
            if (columnIdx < 0)
                return XLError.IncompatibleValue;
            if (columnIdx >= array.Width)
                return XLError.CellReference;

            if (approximateSearchFlag)
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var foundRow = Bisection(array, lookupValue);
                if (foundRow == -1)
                    return XLError.NoValueAvailable;

                return array[foundRow, columnIdx].ToAnyValue();
            }
            else
            {
                // TODO: Implement wildcard search
                for (var rowIndex = 0; rowIndex < array.Height; rowIndex++)
                {
                    var currentValue = array[rowIndex, 0];

                    // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                    var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                    if (comparison == 0)
                        return array[rowIndex, columnIdx].ToAnyValue();
                }

                return XLError.NoValueAvailable;
            }
        }

        private static int Bisection(Array range, ScalarValue lookupValue)
        {
            // Bisection is predicated on a fact that values of the same type are sorted.
            // If they are not, results are unpredictable.
            // Invariants:
            // * Low row has a value that is less or equal than lookup value
            // * High row has a value that is greater than lookup value
            var lowRow = 0;
            var highRow = range.Height - 1;

            lowRow = FindSameTypeRow(range, highRow, 1, lowRow, in lookupValue);
            if (lowRow == -1)
                return -1; // Range doesn't contain even one element of same type

            // Sanity check for unsorted ranges. For bisection to work, lowRow always
            // has to have a value that is less or equal to the lookup value.
            var lowValue = range[lowRow, 0];
            var lowCompare = ScalarValueComparer.SortIgnoreCase.Compare(lowValue, lookupValue);

            // Ensure invariants before main loop. If even lowest value in the range is greater than lookup value,
            // then there can't be any row that matches lookup value/lower.
            if (lowCompare > 0)
                return -1;

            // Since we already know that there is at least one element of same type as lookup value,
            // high row will find something, though it might be same row as lowRow.
            highRow = FindSameTypeRow(range, lowRow, -1, highRow, in lookupValue);

            // Sanity check for unsorted ranges. For bisection to work, highRow always
            // has to have a value that is greater than the lookup value
            var highValue = range[highRow, 0];
            var highCompare = ScalarValueComparer.SortIgnoreCase.Compare(highValue, lookupValue);

            // Ensure invariants before main loop. If the lookup value is greater/equal than
            // the greatest value of the range, it is the result.
            if (highCompare <= 0)
                return highRow;

            // Now we have two borders with actual values and we know the lookup value is less than high and greater/equal to lower
            while (true)
            {
                // The FindMiddle method returns only values [lowRow, highRow)
                // so in each loop it decreases the interval. The lowRow value is
                // the last one checked during search of a middle.
                var middleRow = FindMiddle(range, lowRow, highRow, in lookupValue);

                // A condition for "if an exact match is not found, the next
                // largest value that is less than lookup-value is returned".
                // At this time, lowRow is less than lookup value and highRow
                // is more than lookup value.
                if (middleRow == lowRow)
                    return lowRow;

                var middleValue = range[middleRow, 0];
                var middleCompare = ScalarValueComparer.SortIgnoreCase.Compare(middleValue, lookupValue);

                if (middleCompare <= 0)
                    lowRow = middleRow;
                else
                    highRow = middleRow;
            }
        }

        /// <summary>
        /// Find a row with a value of same type as <paramref name="lookupValue"/>
        /// between values <paramref name="low"/> and <c><paramref name="high"/> - 1</c>.
        /// We know that both <paramref name="low"/> and <paramref name="high"/>
        /// contain value of the same type, so we always get a valid row.
        /// </summary>
        private static int FindMiddle(Array range, int low, int high, in ScalarValue lookupValue)
        {
            Debug.Assert(low < high);
            var middleRow = (low + high) / 2;

            // Since low is < high, it's always possible skip high row for determining middle row
            var higherIndex = FindSameTypeRow(range, high - 1, 1, middleRow, in lookupValue);
            if (higherIndex != -1)
                return higherIndex;

            // We can't skip low like we did for high, because there might be only different type
            // Cells between low row and high row.
            var lowerIndex = FindSameTypeRow(range, low, -1, middleRow, in lookupValue);
            return lowerIndex;
        }

        /// <summary>
        /// Find row index of an element with same type as the lookup value. Go from
        /// <paramref name="startRow"/> to the <paramref name="limitRow"/> by a step
        /// of <paramref name="delta"/>. If there isn't any such row, return <c>-1</c>.
        /// </summary>
        private static int FindSameTypeRow(Array range, int limitRow, int delta, int startRow, in ScalarValue lookupValue)
        {
            // Although the spec says that elements must be sorted in
            // "ascending order", as follows: ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.
            // In reality, comparison ignores elements of the different type than lookupValue.
            // E.g. search for 2.5 in the {"1", 2, "3", #DIV/0!, 3 } will find the second element 2
            // Elements with incompatible type are just skipped.
            int currentRow;
            for (currentRow = startRow; !lookupValue.HaveSameType(range[currentRow, 0]); currentRow += delta)
            {
                // Don't move beyond limitRow
                if (currentRow == limitRow)
                    return -1;
            }

            return currentRow;
        }

        private static AnyValue RowsOrColumns(AnyValue value, bool rows)
        {
            if (value.TryPickArea(out var area, out _))
                return rows ? area.RowSpan : area.ColumnSpan;

            if (value.TryPickArray(out var array))
                return rows ? array.Height : array.Width;

            if (value.TryPickError(out var error))
                return error;

            if (value.IsLogical || value.IsNumber || value.IsText)
                return 1;

            if (value.IsBlank)
                return XLError.IncompatibleValue;

            // Only thing left, if reference has multiple areas
            return XLError.CellReference;
        }
    }
}
