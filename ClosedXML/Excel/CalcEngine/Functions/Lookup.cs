// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
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
            //ce.RegisterFunction("COLUMNS", , Columns); // Returns the number of columns in a reference
            //ce.RegisterFunction("FORMULATEXT", , Formulatext); // Returns the formula at the given reference as text
            //ce.RegisterFunction("GETPIVOTDATA", , Getpivotdata); // Returns data stored in a PivotTable report
            ce.RegisterFunction("HLOOKUP", 3, 4, Hlookup, AllowRange.Only, 1); // Looks in the top row of an array and returns the value of the indicated cell
            ce.RegisterFunction("HYPERLINK", 1, 2, Adapt(Hyperlink), FunctionFlags.Scalar | FunctionFlags.SideEffect); // Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
            ce.RegisterFunction("INDEX", 2, 4, Index, AllowRange.Only, 0, 1); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            ce.RegisterFunction("MATCH", 2, 3, Match, AllowRange.Only, 1); // Looks up values in a reference or array
            //ce.RegisterFunction("OFFSET", , Offset); // Returns a reference offset from a given reference
            ce.RegisterFunction("ROW", 0, 1, Row, FunctionFlags.Range, AllowRange.All); // Returns the row number of a reference
            //ce.RegisterFunction("ROWS", , Rows); // Returns the number of rows in a reference
            //ce.RegisterFunction("RTD", , Rtd); // Retrieves real-time data from a program that supports COM automation
            //ce.RegisterFunction("TRANSPOSE", , Transpose); // Returns the transpose of an array
            ce.RegisterFunction("VLOOKUP", 3, 4, Vlookup, AllowRange.Only, 1); // Looks in the first column of an array and moves across the row to return the value of a cell
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

        private static bool TryExtractRange(Expression expression, out IXLRange range, out XLError calculationErrorType)
        {
            range = null;
            calculationErrorType = default;

            if (!(expression is XObjectExpression objectExpression))
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            if (!(objectExpression.Value is CellRangeReference cellRangeReference))
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            range = cellRangeReference.Range;
            return true;
        }

        private static object Hlookup(List<Expression> p)
        {
            var lookup_value = p[0];

            if (!TryExtractRange(p[1], out var range, out var error))
                return error;

            var row_index_num = (int)p[2];
            var range_lookup = p.Count < 4
                               || p[3] is EmptyValueExpression
                               || (bool)(p[3]);

            if (row_index_num < 1)
                return XLError.CellReference;

            if (row_index_num > range.RowCount())
                return XLError.CellReference;

            IXLRangeColumn matching_column;
            matching_column = range.FindColumn(c => !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) == 0);
            if (range_lookup && matching_column == null)
            {
                var first_column = range.FirstColumn().ColumnNumber();
                var number_of_columns_in_range = range.ColumnsUsed().Count();

                matching_column = range.FindColumn(c =>
                {
                    var column_index_in_range = c.ColumnNumber() - first_column + 1;
                    if (column_index_in_range < number_of_columns_in_range && !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) <= 0 && !c.ColumnRight().Cell(1).IsEmpty() && new Expression(c.ColumnRight().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (column_index_in_range == number_of_columns_in_range && !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_column == null)
                return XLError.NoValueAvailable;

            return matching_column
                .Cell(row_index_num)
                .Value;
        }

        private static AnyValue Hyperlink(CalcContext ctx, string linkLocation, ScalarValue? friendlyName)
        {
            var link = new XLHyperlink(linkLocation);
            var cell = ctx.Worksheet.Cell(ctx.FormulaAddress);
            cell.SetHyperlink(link);

            return friendlyName?.ToAnyValue() ?? linkLocation;
        }

        private static object Index(List<Expression> p)
        {
            // This is one of the few functions that is "overloaded"
            if (!TryExtractRange(p[0], out var range, out var error))
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

            if (!TryExtractRange(p[1], out var range, out var error))
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

        private static object Vlookup(List<Expression> p)
        {
            var lookup_value = p[0];

            if (!TryExtractRange(p[1], out var range, out var error))
                return error;

            var col_index_num = (int)p[2];
            var range_lookup = p.Count < 4
                               || p[3] is EmptyValueExpression
                               || (bool)(p[3]);

            if (col_index_num < 1)
                return XLError.CellReference;

            if (col_index_num > range.ColumnCount())
                return XLError.CellReference;

            IXLRangeRow matching_row;
            try
            {
                matching_row = range.FindRow(r => !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) == 0);
            }
            catch (Exception)
            {
                return XLError.NoValueAvailable;
            }
            if (range_lookup && matching_row == null)
            {
                var first_row = range.FirstRow().RowNumber();
                var number_of_rows_in_range = range.RowsUsed().Count();

                matching_row = range.FindRow(r =>
                {
                    var row_index_in_range = r.RowNumber() - first_row + 1;
                    if (row_index_in_range < number_of_rows_in_range && !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) <= 0 && !r.RowBelow().Cell(1).IsEmpty() && new Expression(r.RowBelow().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (row_index_in_range == number_of_rows_in_range && !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_row == null)
                return XLError.NoValueAvailable;

            return matching_row
                .Cell(col_index_num)
                .Value;
        }
    }
}
