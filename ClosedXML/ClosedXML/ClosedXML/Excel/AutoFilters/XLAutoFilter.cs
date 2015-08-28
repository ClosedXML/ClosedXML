using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;
  
    public class XLAutoFilter : IXLBaseAutoFilter, IXLAutoFilter
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
            Enabled = false;
            Filters.Clear();
            foreach (IXLRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
                row.WorksheetRow().Unhide();
            return this;
        }

        public XLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
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


        /// <summary>
        /// Contraty to individual column filtering, this method applies all filters aggregated results at once on all
        /// rows within the AutoFilter's range.
        /// </summary>
        public void ReapplyAllFilter()
        {
          var _autoFilter = this;
          var ws = _autoFilter.Range.Worksheet as XLWorksheet;
          ws.SuspendEvents();
          var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
          foreach (IXLRangeRow _row in rows)
          {
            Boolean visible = true;

            //Go through on each filter for each column
            foreach (var filter in _autoFilter.Filters)
            {
              var _column = filter.Key;
              var _cell = _row.Cell(_column);
              Boolean isText = _cell.DataType == XLCellValues.Text;

              //Set the default filterMatch OR => false, AND => true;
              Boolean filterMatch = true;
              Boolean firstTime = true;

              //Go though on each item condition in each filter
              foreach (var filterItem in (List<XLFilter>)filter.Value)
              {
                //Initialize filterMatch for the first time:
                if (firstTime)
                {
                  firstTime = false;
                  filterMatch = (filterItem.Connector == XLConnector.And);
                }

                //match the value for each filterItem
                Boolean itemMatch = isText
                           ? filterItem.Condition(_cell.GetString())
                           : _cell.DataType == XLCellValues.Number &&
                               filterItem.Condition(_cell.GetDouble());

                //Summarize by the operator
                if (filterItem.Connector == XLConnector.And)
                  filterMatch = filterMatch && itemMatch;
                else
                  filterMatch = filterMatch || itemMatch;
              }

              //Summarize filterMatch to adjust row visibility
              visible = visible && filterMatch;

              //Break out quicker for already hidden rows
              if (!visible)
                break;
            }

            if (visible)
              _row.WorksheetRow().Unhide();
            else
              _row.WorksheetRow().Hide();
          }
          ws.ResumeEvents();
        }
     }
  }