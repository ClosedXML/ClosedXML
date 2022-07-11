// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Lookup
    {
        public static void Register(CalcEngine ce)
        {
            //ce.RegisterFunction("ADDRESS", , Address); // Returns a reference as text to a single cell in a worksheet
            //ce.RegisterFunction("AREAS", , Areas); // Returns the number of areas in a reference
            //ce.RegisterFunction("CHOOSE", , Choose); // Chooses a value from a list of values
            //ce.RegisterFunction("COLUMN", , Column); // Returns the column number of a reference
            //ce.RegisterFunction("COLUMNS", , Columns); // Returns the number of columns in a reference
            //ce.RegisterFunction("FORMULATEXT", , Formulatext); // Returns the formula at the given reference as text
            //ce.RegisterFunction("GETPIVOTDATA", , Getpivotdata); // Returns data stored in a PivotTable report
            ce.RegisterFunction("HLOOKUP", 3, 4, Hlookup); // Looks in the top row of an array and returns the value of the indicated cell
            ce.RegisterFunction("HYPERLINK", 1, 2, Hyperlink); // Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
            ce.RegisterFunction("INDEX", 2, 4, Index); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            ce.RegisterFunction("MATCH", 2, 3, Match); // Looks up values in a reference or array
            //ce.RegisterFunction("OFFSET", , Offset); // Returns a reference offset from a given reference
            //ce.RegisterFunction("ROW", , Row); // Returns the row number of a reference
            //ce.RegisterFunction("ROWS", , Rows); // Returns the number of rows in a reference
            //ce.RegisterFunction("RTD", , Rtd); // Retrieves real-time data from a program that supports COM automation
            //ce.RegisterFunction("TRANSPOSE", , Transpose); // Returns the transpose of an array
            ce.RegisterFunction("VLOOKUP", 3, 4, Vlookup); // Looks in the first column of an array and moves across the row to return the value of a cell
        }

        private static IXLRange ExtractRange(Expression expression)
        {
            if (!(expression is XObjectExpression objectExpression))
                throw new NoValueAvailableException("Parameter has to be a valid range");

            if (!(objectExpression.Value is CellRangeReference cellRangeReference))
                throw new NoValueAvailableException("lookup_array has to be a range");

            var range = cellRangeReference.Range;
            return range;
        }

        private static object Hlookup(List<Expression> p)
        {
            var lookup_value = p[0];
            var range = ExtractRange(p[1]);
            var row_index_num = (int)p[2];
            var range_lookup = p.Count < 4
                               || p[3] is EmptyValueExpression
                               || (bool)(p[3]);

            if (row_index_num < 1)
                throw new CellReferenceException("Row index has to be positive");

            if (row_index_num > range.RowCount())
                throw new CellReferenceException("Row index has to be positive");

            IXLRangeColumn matching_column;
            matching_column = range.FindColumn(c => !c.Cell(1).IsEmpty() && new ScalarNode(c.Cell(1).Value).CompareTo(lookup_value) == 0);
            if (range_lookup && matching_column == null)
            {
                var first_column = range.FirstColumn().ColumnNumber();
                var number_of_columns_in_range = range.ColumnsUsed().Count();

                matching_column = range.FindColumn(c =>
                {
                    var column_index_in_range = c.ColumnNumber() - first_column + 1;
                    if (column_index_in_range < number_of_columns_in_range && !c.Cell(1).IsEmpty() && new ScalarNode(c.Cell(1).Value).CompareTo(lookup_value) <= 0 && !c.ColumnRight().Cell(1).IsEmpty() && new ScalarNode(c.ColumnRight().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (column_index_in_range == number_of_columns_in_range && !c.Cell(1).IsEmpty() && new ScalarNode(c.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_column == null)
                throw new NoValueAvailableException("No matches found.");

            return matching_column
                .Cell(row_index_num)
                .Value;
        }

        private static object Hyperlink(List<Expression> p)
        {
            String address = p[0];
            String toolTip = p.Count == 2 ? p[1] : String.Empty;
            return new XLHyperlink(address, toolTip);
        }

        private static object Index(List<Expression> p)
        {
            // This is one of the few functions that is "overloaded"
            var range = ExtractRange(p[0]);

            if (range.ColumnCount() > 1 && range.RowCount() > 1)
            {
                var row_num = (int)p[1];
                var column_num = (int)p[2];

                if (row_num > range.RowCount())
                    throw new CellReferenceException("Out of bound row number");

                if (column_num > range.ColumnCount())
                    throw new CellReferenceException("Out of bound column number");

                return range.Row(row_num).Cell(column_num).Value;
            }
            else if (p.Count == 2)
            {
                var cellOffset = (int)p[1];
                if (cellOffset > range.RowCount() * range.ColumnCount())
                    throw new CellReferenceException();

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
                    throw new CellReferenceException();

                if (!rangeIsRow && column_num > 1)
                    throw new CellReferenceException();

                if (row_num > range.RowCount())
                    throw new CellReferenceException("Out of bound row number");

                if (column_num > range.ColumnCount())
                    throw new CellReferenceException("Out of bound column number");

                return range.Row(row_num).Cell(column_num).Value;
            }
        }

        private static object Match(List<Expression> p)
        {
            var lookup_value = p[0];
            var range = ExtractRange(p[1]);
            int match_type = 1;
            if (p.Count > 2)
                match_type = Math.Sign((int)p[2]);

            if (range.ColumnCount() != 1 && range.RowCount() != 1)
                throw new CellValueException("Range has to be 1-dimensional");

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
                    throw new NoValueAvailableException("Invalid match_type");
            }

            IXLCell foundCell = null;

            if (match_type == 0)
                foundCell = range
                    .CellsUsed(XLCellsUsedOptions.Contents, c => lookupPredicate.Invoke(new ScalarNode(c.Value).CompareTo(lookup_value)))
                    .FirstOrDefault();
            else
            {
                object previousValue = null;
                foundCell = range
                    .CellsUsed(XLCellsUsedOptions.Contents)
                    .TakeWhile(c =>
                    {
                        var currentCellExpression = new ScalarNode(c.Value);

                        if (previousValue != null)
                        {
                            // When match_type != 0, we have to assume that the order of the items being search is ascending or descending
                            var previousValueExpression = new ScalarNode(previousValue);
                            if (!lookupPredicate.Invoke(previousValueExpression.CompareTo(currentCellExpression)))
                                return false;
                        }

                        previousValue = c.Value;

                        return lookupPredicate.Invoke(currentCellExpression.CompareTo(lookup_value));
                    })
                    .LastOrDefault();
            }

            if (foundCell == null)
                throw new NoValueAvailableException();

            var firstCell = range.FirstCell();

            return (foundCell.Address.ColumnNumber - firstCell.Address.ColumnNumber + 1) * (foundCell.Address.RowNumber - firstCell.Address.RowNumber + 1);
        }

        private static object Vlookup(List<Expression> p)
        {
            var lookup_value = p[0];
            var range = ExtractRange(p[1]);
            var col_index_num = (int)p[2];
            var range_lookup = p.Count < 4
                               || p[3] is EmptyValueExpression
                               || (bool)(p[3]);

            if (col_index_num < 1)
                throw new CellReferenceException("Column index has to be positive");

            if (col_index_num > range.ColumnCount())
                throw new CellReferenceException("Colum index must be smaller or equal to the number of columns in the table array");

            IXLRangeRow matching_row;
            try
            {
                matching_row = range.FindRow(r => !r.Cell(1).IsEmpty() && new ScalarNode(r.Cell(1).Value).CompareTo(lookup_value) == 0);
            }
            catch (Exception ex)
            {
                throw new NoValueAvailableException("No matches found", ex);
            }
            if (range_lookup && matching_row == null)
            {
                var first_row = range.FirstRow().RowNumber();
                var number_of_rows_in_range = range.RowsUsed().Count();

                matching_row = range.FindRow(r =>
                {
                    var row_index_in_range = r.RowNumber() - first_row + 1;
                    if (row_index_in_range < number_of_rows_in_range && !r.Cell(1).IsEmpty() && new ScalarNode(r.Cell(1).Value).CompareTo(lookup_value) <= 0 && !r.RowBelow().Cell(1).IsEmpty() && new ScalarNode(r.RowBelow().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (row_index_in_range == number_of_rows_in_range && !r.Cell(1).IsEmpty() && new ScalarNode(r.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_row == null)
                throw new NoValueAvailableException("No matches found.");

            return matching_row
                .Cell(col_index_num)
                .Value;
        }
    }
}
