using System;
using System.Linq;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLAutoFilter: IXLBaseAutoFilter, IXLAutoFilter
    {
        public XLAutoFilter()
        {
            Filters = new Dictionary<int, List<XLFilter>>();
        }

        public Boolean Enabled { get; set; }
        public IXLRange Range { get; set; }
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
            foreach (var row in Range.Rows().Where(r => r.RowNumber() > 1))
            {
                row.WorksheetRow().Unhide();
            }
            return this;
        }

        IXLBaseAutoFilter IXLBaseAutoFilter.Clear()
        {
            return Clear();
        }

        IXLBaseAutoFilter IXLBaseAutoFilter.Set(IXLRangeBase range)
        {
            return Set(range);
        }

        public XLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            
            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            if (Enabled)
            {
                foreach (var row in Range.Rows().Where(r => r.RowNumber() > 1))
                {
                    row.WorksheetRow().Unhide();
                }
                foreach (var kp in Filters)
                {
                    Boolean firstFilter = true;
                    foreach (XLFilter filter in kp.Value)
                    {
                        Boolean isText = filter.Value is String;
                        foreach (var row in Range.Rows().Where(r => r.RowNumber() > 1))
                        {
                            Boolean match = isText ? filter.Condition(row.Cell(kp.Key).GetString()) : row.Cell(kp.Key).DataType == XLCellValues.Number && filter.Condition(row.Cell(kp.Key).GetDouble());
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
                Enabled = true;
            }

            return this;
        }

        IXLAutoFilter IXLAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        IXLBaseAutoFilter IXLBaseAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public Boolean Sorted { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public Int32 SortColumn { get; set; }

        public Dictionary<Int32, List<XLFilter>> Filters { get; private set; }
        
        //List<IXLFilter> Filters { get; }
        //List<IXLCustomFilter> CustomFilters { get; }
        //Boolean Sorted { get; }
        //Int32 SortColumn { get; }
        public IXLFilterColumn Column(String column)
        {
            return Column(ExcelHelper.GetColumnNumberFromLetter(column));
        }
        Dictionary<Int32, XLFilterColumn> _columns = new Dictionary<int, XLFilterColumn>();
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
    }
}