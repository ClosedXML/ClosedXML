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
            ce.RegisterFunction("INDEX", 2, 4, AdaptIndex(Index), FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.Only, 0); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            ce.RegisterFunction("MATCH", 2, 3, AdaptMatch(Match), FunctionFlags.Range, AllowRange.Only, 1); // Looks up values in a reference or array
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

        public static AnyValue Index(CalcContext ctx, AnyValue value, List<int> p)
        {
            var areaNumber = p.Count > 2 ? p[2] : 1;
            if (areaNumber < 1)
                return XLError.IncompatibleValue;

            if (!value.IsReference && areaNumber > 1)
                return XLError.CellReference;

            // There must be two paths, one for array and one for reference. Reference path
            // must return reference, so it behaves correctly with implicit intersection.
            OneOf<XLRangeAddress, Array> data;
            if (value.TryPickScalar(out var scalar, out var collection))
            {
                if (scalar.IsBlank)
                    return XLError.IncompatibleValue;

                data = new ScalarArray(scalar, 1, 1);
            }
            else if (collection.TryPickT0(out var valueArray, out var reference))
            {
                data = valueArray;
            }
            else
            {
                if (areaNumber > reference.Areas.Count)
                    return XLError.CellReference;

                data = reference.Areas[areaNumber - 1];
            }

            var width = data.Match(static area => area.ColumnSpan, static array => array.Width);
            var height = data.Match(static area => area.RowSpan, static array => array.Height);

            var rowNumber = 0;
            var colNumber = 0;
            if (p.Count == 1)
            {
                if (width == 1)
                    rowNumber = p[0];

                if (height == 1)
                    colNumber = p[0];
            }

            if (p.Count >= 2)
            {
                rowNumber = p[0];
                colNumber = p[1];
            }

            // Check the bounded values
            if (rowNumber < 0 || colNumber < 0)
                return XLError.IncompatibleValue;

            if (rowNumber > height || colNumber > width)
                return XLError.CellReference;

            return data.TryPickT0(out var area, out var array)
                ? IndexArea(area, rowNumber, colNumber)
                : IndexArray(array, rowNumber, colNumber);

            static Reference IndexArea(XLRangeAddress area, int rowNumber, int colNumber)
            {
                // Return whole area
                if (rowNumber == 0 && colNumber == 0)
                    return new Reference(area);

                // Return one column at colNumber
                if (rowNumber == 0)
                {
                    var topCell = new XLAddress(area.Worksheet, area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber + colNumber - 1, true, true);
                    var bottomCell = new XLAddress(area.Worksheet, area.LastAddress.RowNumber, area.FirstAddress.ColumnNumber + colNumber - 1, true, true);
                    return new Reference(new XLRangeAddress(topCell, bottomCell));
                }

                // Return one row at rowNumber
                if (colNumber == 0)
                {
                    var leftCell = new XLAddress(area.Worksheet, area.FirstAddress.RowNumber + rowNumber - 1, area.FirstAddress.ColumnNumber, true, true);
                    var rightCell = new XLAddress(area.Worksheet, area.FirstAddress.RowNumber + rowNumber - 1, area.LastAddress.ColumnNumber, true, true);
                    return new Reference(new XLRangeAddress(leftCell, rightCell));
                }

                // Return single cell reference.
                var areaCorner = area.FirstAddress;
                var cellAddress = new XLAddress(area.Worksheet, areaCorner.RowNumber + rowNumber - 1, areaCorner.ColumnNumber + colNumber - 1, true, true);
                return new Reference(new XLRangeAddress(cellAddress, cellAddress));
            }

            static AnyValue IndexArray(Array array, int rowNumber, int colNumber)
            {
                // Return whole array
                if (rowNumber == 0 && colNumber == 0)
                    return array;

                // Return one column at colNumber
                if (rowNumber == 0)
                    return new SlicedArray(array, 0, array.Height, colNumber - 1, 1);

                // Return one row at rowNumber
                if (colNumber == 0)
                    return new SlicedArray(array, rowNumber - 1, 1, 0, array.Width);

                // Return single value
                return array[rowNumber - 1, colNumber - 1].ToAnyValue();
            }
        }

        private static ScalarValue Match(CalcContext ctx, ScalarValue target, AnyValue lookupArray, int matchType)
        {
            if (target.IsBlank)
                return XLError.NoValueAvailable;

            if (target.TryPickError(out var error))
                return error;

            if (!lookupArray.TryPickCollectionArray(out var array, ctx))
                return XLError.NoValueAvailable;

            // Match only supports arrays with one row or one column.
            // Normalize to an array with one column in both cases.
            if (array.Height == 1 && array.Width > 1)
                array = new TransposedArray(array);

            if (array.Width != 1)
                return XLError.NoValueAvailable;

            var index = matchType switch
            {
                < 0 => MatchDescending(target, array, ScalarValueComparer.SortIgnoreCase),
                0 => MatchUnsorted(target, array, ctx),
                > 0 => MatchAscending(target, array, ScalarValueComparer.SortIgnoreCase),
            };

            if (index < 0)
                return XLError.NoValueAvailable;

            return index + 1;

            static int MatchAscending(ScalarValue target, Array data, IComparer<ScalarValue> comparer)
            {
                var index = Bisection(target, data, comparer);
                if (index == -1)
                    return index;

                // When there are multiple same elements, return position of the last one
                while (index < data.Height - 1 && comparer.Compare(data[index + 1, 0], data[index, 0]) == 0)
                    index++;

                return index;
            }

            static int MatchUnsorted(ScalarValue target, Array data, CalcContext ctx)
            {
                var criteria = Criteria.Create(target, ctx.Culture);
                for (var i = 0; i < data.Height; ++i)
                {
                    var value = data[i, 0];
                    if (target.HaveSameType(value) && criteria.Match(value))
                        return i;
                }

                return -1;
            }

            static int MatchDescending(ScalarValue target, Array data, IComparer<ScalarValue> comparer)
            {
                // Data should be in ascending order, but Excel doesn't use bisection.
                var found = -1;
                for (var i = 0; i < data.Height; ++i)
                {
                    // Skip elements with different type
                    var value = data[i, 0];
                    while (!value.HaveSameType(target))
                    {
                        if (i == data.Height - 1)
                            return found;

                        value = data[++i, 0];
                    }

                    var compare = comparer.Compare(target, value);
                    if (compare == 0)
                        return i;

                    if (compare > 0) // target > value
                        return found;

                    // value > target, so there might an exact match later
                    found = i;
                }

                return found;
            }
        }

        /// <summary>
        /// Find index of the greatest element smaller or equal to the <paramref name="target"/>.
        /// </summary>
        /// <param name="target">Value to look for.</param>
        /// <param name="data">Data in ascending order.</param>
        /// <param name="comparer">A comparator for comparing two values.</param>
        /// <returns>Index of found element. If the <paramref name="data"/> contains
        ///   a sequence of <paramref name="target"/> values, it can be index of any of them.
        /// </returns>
        private static int Bisection(ScalarValue target, Array data, IComparer<ScalarValue> comparer)
        {
            // This should match Excel logic perfectly. Make sure to do some fuzzy testing when changing the code.
            var low = 0;
            var high = data.Height - 1;
            while (low < high)
            {
                var (middle, compare) = FindMiddleAbove(low, high, target, data, comparer);

                if (compare == 0)
                    return middle;

                // target < value
                if (compare < 0)
                    high = Math.Max(low, middle - 1);

                // target > value
                if (compare > 0)
                    low = Math.Min(high, middle + 1);
            }

            // Final index might point to an element greater than the lookup
            // (e.g. { 1, 2 } with lookup 1.5). The data should be ascending,
            // so just go in the expected order.
            for (var i = low; i >= 0; --i)
            {
                var compare = comparer.Compare(data[i, 0], target);
                if (compare <= 0) // data[i] <= target
                    return i;
            }

            return -1;

            static (int Middle, int Comparison) FindMiddleAbove(int low, int high, ScalarValue target, Array data, IComparer<ScalarValue> comparer)
            {
                var initial = (low + high) / 2;
                var middle = initial;
                while (middle <= high)
                {
                    if (data[middle, 0].HaveSameType(target))
                        return (middle, comparer.Compare(target, data[middle, 0]));

                    middle++;
                }

                // There is nothing left in the higher half. Target must be in the lower half.
                return (initial, -1);
            }
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
