// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLAutoFilter : IXLAutoFilter
    {
        private readonly Dictionary<Int32, XLFilterColumn> _columns = new Dictionary<int, XLFilterColumn>();

        public XLAutoFilter()
        {
            Filters = new Dictionary<int, List<XLFilter>>();
        }

        public Dictionary<Int32, List<XLFilter>> Filters { get; private set; }

        #region IXLAutoFilter Members

        public Boolean Enabled { get; set; }
        public IEnumerable<IXLRangeRow> HiddenRows { get => Range.Rows(r => r.WorksheetRow().IsHidden); }
        public IXLRange Range { get; set; }
        public Int32 SortColumn { get; set; }
        public Boolean Sorted { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public IEnumerable<IXLRangeRow> VisibleRows { get => Range.Rows(r => !r.WorksheetRow().IsHidden); }

        IXLAutoFilter IXLAutoFilter.Clear()
        {
            return Clear();
        }

        public IXLFilterColumn Column(String column)
        {
            var columnNumber = XLHelper.GetColumnNumberFromLetter(column);
            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(column), "Column '" + column + "' is outside the allowed column range.");

            return Column(columnNumber);
        }

        public IXLFilterColumn Column(Int32 column)
        {
            if (column < 1 || column > XLHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(column), "Column " + column + " is outside the allowed column range.");

            if (!_columns.TryGetValue(column, out XLFilterColumn filterColumn))
            {
                filterColumn = new XLFilterColumn(this, column);
                _columns.Add(column, filterColumn);
            }

            return filterColumn;
        }

        public IXLAutoFilter Reapply()
        {
            var ws = Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();

            // Recalculate shown / hidden rows
            var rows = Range.Rows(2, Range.RowCount());
            rows.ForEach(row =>
                row.WorksheetRow().Unhide()
            );

            foreach (IXLRangeRow row in rows)
            {
                var rowMatch = true;

                foreach (var columnIndex in Filters.Keys)
                {
                    var columnFilters = Filters[columnIndex];

                    var columnFilterMatch = true;

                    // If the first filter is an 'Or', we need to fudge the initial condition
                    if (columnFilters.Count > 0 && columnFilters.First().Connector == XLConnector.Or)
                    {
                        columnFilterMatch = false;
                    }

                    foreach (var filter in columnFilters)
                    {
                        var condition = filter.Condition;
                        var isText = filter.Value is String;
                        var isDateTime = filter.Value is DateTime;

                        Boolean filterMatch;

                        if (isText)
                            filterMatch = condition(row.Cell(columnIndex).GetFormattedString());
                        else if (isDateTime)
                            filterMatch = row.Cell(columnIndex).DataType == XLDataType.DateTime &&
                                    condition(row.Cell(columnIndex).GetDateTime());
                        else
                            filterMatch = row.Cell(columnIndex).DataType == XLDataType.Number &&
                                    condition(row.Cell(columnIndex).GetDouble());

                        if (filter.Connector == XLConnector.And)
                        {
                            columnFilterMatch &= filterMatch;
                            if (!columnFilterMatch) break;
                        }
                        else
                        {
                            columnFilterMatch |= filterMatch;
                            if (columnFilterMatch) break;
                        }
                    }

                    rowMatch &= columnFilterMatch;

                    if (!rowMatch) break;
                }

                if (!rowMatch) row.WorksheetRow().Hide();
            }

            ws.ResumeEvents();
            return this;
        }

        IXLAutoFilter IXLAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase,
                                                                                                         Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        #endregion IXLAutoFilter Members

        public XLAutoFilter Clear()
        {
            if (!Enabled) return this;

            Enabled = false;
            Filters.Clear();
            foreach (IXLRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
                row.WorksheetRow().Unhide();
            return this;
        }

        public XLAutoFilter Set(IXLRangeBase range)
        {
            Range = range.AsRange();
            Enabled = true;
            return this;
        }

        public XLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            if (!Enabled)
                throw new InvalidOperationException("Filter has not been enabled.");

            var ws = Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
                                                                              ignoreBlanks);

            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            ws.ResumeEvents();

            Reapply();

            return this;
        }
    }
}
