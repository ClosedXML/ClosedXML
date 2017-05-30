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
            ce.RegisterFunction("HLOOKUP", 4, Hlookup); // Looks in the top row of an array and returns the value of the indicated cell
            //ce.RegisterFunction("HYPERLINK", , Hyperlink); // Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
            //ce.RegisterFunction("INDEX", , Index); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            //ce.RegisterFunction("MATCH", , Match); // Looks up values in a reference or array
            //ce.RegisterFunction("OFFSET", , Offset); // Returns a reference offset from a given reference
            //ce.RegisterFunction("ROW", , Row); // Returns the row number of a reference
            //ce.RegisterFunction("ROWS", , Rows); // Returns the number of rows in a reference
            //ce.RegisterFunction("RTD", , Rtd); // Retrieves real-time data from a program that supports COM automation
            //ce.RegisterFunction("TRANSPOSE", , Transpose); // Returns the transpose of an array
            ce.RegisterFunction("VLOOKUP", 4, Vlookup); // Looks in the first column of an array and moves across the row to return the value of a cell
        }

        private static object Hlookup(List<Expression> p)
        {
            var lookup_value = p[0];

            var table_array = p[1] as XObjectExpression;
            var range_reference = table_array.Value as CellRangeReference;
            var range = range_reference.Range;

            var row_index_num = (int)(p[2]);
            var range_lookup = p.Count < 4 || (bool)(p[3]);

            if (table_array == null || range_reference == null)
                throw new ApplicationException("table_array has to be a range");

            if (row_index_num < 1)
                throw new ApplicationException("col_index_num has to be positive");

            if (row_index_num > range.RowCount())
                throw new ApplicationException("col_index_num must be smaller or equal to the number of rows in the table array");

            IXLRangeColumn matching_column;
            matching_column = range.FindColumn(c => !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) == 0);
            if (range_lookup && matching_column == null)
            {
                var first_column = range.FirstColumn().ColumnNumber();
                matching_column = range.FindColumn(c =>
                {
                    var column_index_in_range = c.ColumnNumber() - first_column + 1;
                    if (column_index_in_range < range.ColumnsUsed().Count() && !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) <= 0 && !c.ColumnRight().Cell(1).IsEmpty() && new Expression(c.ColumnRight().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (column_index_in_range == range.ColumnsUsed().Count() && !c.Cell(1).IsEmpty() && new Expression(c.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_column == null)
                throw new ApplicationException("No matches found.");

            return matching_column
                .Cell(row_index_num)
                .Value;
        }

        private static object Vlookup(List<Expression> p)
        {
            var lookup_value = p[0];

            var table_array = p[1] as XObjectExpression;
            var range_reference = table_array.Value as CellRangeReference;
            var range = range_reference.Range;

            var col_index_num = (int)(p[2]);
            var range_lookup = p.Count < 4 || (bool)(p[3]);

            if (table_array == null || range_reference == null)
                throw new ApplicationException("table_array has to be a range");

            if (col_index_num < 1)
                throw new ApplicationException("col_index_num has to be positive");

            if (col_index_num > range.ColumnCount())
                throw new ApplicationException("col_index_num must be smaller or equal to the number of columns in the table array");

            IXLRangeRow matching_row;
            matching_row = range.FindRow(r => !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) == 0);
            if (range_lookup && matching_row == null)
            {
                var first_row = range.FirstRow().RowNumber();
                matching_row = range.FindRow(r =>
                {
                    var row_index_in_range = r.RowNumber() - first_row + 1;
                    if (row_index_in_range < range.RowsUsed().Count() && !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) <= 0 && !r.RowBelow().Cell(1).IsEmpty() && new Expression(r.RowBelow().Cell(1).Value).CompareTo(lookup_value) > 0)
                        return true;
                    else if (row_index_in_range == range.RowsUsed().Count() && !r.Cell(1).IsEmpty() && new Expression(r.Cell(1).Value).CompareTo(lookup_value) <= 0)
                        return true;
                    else
                        return false;
                });
            }

            if (matching_row == null)
                throw new ApplicationException("No matches found.");

            return matching_row
                .Cell(col_index_num)
                .Value;
        }
    }
}
