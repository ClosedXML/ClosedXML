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
            foreach (IXLRangeRow row in rows)
                row.WorksheetRow().Unhide();

            foreach (var kp in Filters)
            {
                Boolean firstFilter = true;
                foreach (XLFilter filter in kp.Value)
                {
                    var condition = filter.Condition;
                    var isText = filter.Value is String;
                    var isDateTime = filter.Value is DateTime;

                    foreach (IXLRangeRow row in rows)
                    {
                        //TODO : clean up filter matching - it's done in different place
                        Boolean match;

                        if (isText)
                            match = condition(row.Cell(kp.Key).GetFormattedString());
                        else if (isDateTime)
                            match = row.Cell(kp.Key).DataType == XLDataType.DateTime &&
                                    condition(row.Cell(kp.Key).GetDateTime());
                        else
                            match = row.Cell(kp.Key).DataType == XLDataType.Number &&
                                    condition(row.Cell(kp.Key).GetDouble());

                        if (firstFilter)
                        {
                            if (match)
                                row.WorksheetRow().Unhide();
                            else
                                row.WorksheetRow().Hide();
                        }
                        else
                        {
                            if (filter.Connector == XLConnector.And)
                            {
                                if (!row.WorksheetRow().IsHidden)
                                {
                                    if (match)
                                        row.WorksheetRow().Unhide();
                                    else
                                        row.WorksheetRow().Hide();
                                }
                            }
                            else if (match)
                                row.WorksheetRow().Unhide();
                        }
                    }

                    firstFilter = false;
                }
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
