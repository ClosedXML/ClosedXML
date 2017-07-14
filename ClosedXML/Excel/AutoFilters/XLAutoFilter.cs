using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLAutoFilter : IXLBaseAutoFilter, IXLAutoFilter
    {
        private readonly Dictionary<Int32, XLFilterColumn> _columns = new Dictionary<int, XLFilterColumn>();

        public XLAutoFilter()
        {
            Filters = new Dictionary<int, List<XLFilter>>();
        }

        public Dictionary<Int32, List<XLFilter>> Filters { get; private set; }

        #region IXLAutoFilter Members

        IXLAutoFilter IXLAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase,
                                         Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public void Dispose()
        {
            if (Range != null)
                Range.Dispose();
        }

        #endregion

        #region IXLBaseAutoFilter Members

        public Boolean Enabled { get; set; }
        public IXLRange Range { get; set; }

        IXLBaseAutoFilter IXLBaseAutoFilter.Clear()
        {
            return Clear();
        }

        IXLBaseAutoFilter IXLBaseAutoFilter.Set(IXLRangeBase range)
        {
            return Set(range);
        }

        IXLBaseAutoFilter IXLBaseAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase,
                                                 Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public Boolean Sorted { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public Int32 SortColumn { get; set; }

        public IXLFilterColumn Column(String column)
        {
            return Column(XLHelper.GetColumnNumberFromLetter(column));
        }

        public IXLFilterColumn Column(Int32 column)
        {
            XLFilterColumn filterColumn;
            if (!_columns.TryGetValue(column, out filterColumn))
            {
                filterColumn = new XLFilterColumn(this, column);
                _columns.Add(column, filterColumn);
            }

            return filterColumn;
        }

        #endregion

        public XLAutoFilter Set(IXLRangeBase range)
        {
            Range = range.AsRange();
            Enabled = true;
            return this;
        }

        public XLAutoFilter Clear()
        {
            if (!Enabled) return this;

            Enabled = false;
            Filters.Clear();
            foreach (IXLRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
                row.WorksheetRow().Unhide();
            return this;
        }

        public XLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            if (!Enabled)
                throw new ApplicationException("Filter has not been enabled.");

            var ws = Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
                                                                              ignoreBlanks);

            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            if (Enabled)
            {
                using (var rows = Range.Rows(2, Range.RowCount()))
                {
                    foreach (IXLRangeRow row in rows)
                        row.WorksheetRow().Unhide();
                }

                foreach (KeyValuePair<int, List<XLFilter>> kp in Filters)
                {
                    Boolean firstFilter = true;
                    foreach (XLFilter filter in kp.Value)
                    {
                        Boolean isText = filter.Value is String;
                        using (var rows = Range.Rows(2, Range.RowCount()))
                        {
                            foreach (IXLRangeRow row in rows)
                            {
                                Boolean match = isText
                                                    ? filter.Condition(row.Cell(kp.Key).GetString())
                                                    : row.Cell(kp.Key).DataType == XLCellValues.Number &&
                                                      filter.Condition(row.Cell(kp.Key).GetDouble());
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
                }
            }
            ws.ResumeEvents();
            return this;
        }
    }
}