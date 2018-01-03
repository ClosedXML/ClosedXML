using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{SourceName}")]
    public class XLPivotField : IXLPivotField
    {
        public XLPivotField(string sourceName)
        {
            SourceName = sourceName;
            Subtotals = new List<XLSubtotalFunction>();
            SelectedValues = new List<Object>();
            SortType = XLPivotSortType.Default;
        }

        public String SourceName { get; private set; }
        public String CustomName { get; set; }

        public IXLPivotField SetCustomName(String value) { CustomName = value; return this; }

        public List<XLSubtotalFunction> Subtotals { get; private set; }

        public IXLPivotField AddSubtotal(XLSubtotalFunction value) { Subtotals.Add(value); return this; }

        public Boolean IncludeNewItemsInFilter { get; set; }

        public IXLPivotField SetIncludeNewItemsInFilter() { IncludeNewItemsInFilter = true; return this; }

        public IXLPivotField SetIncludeNewItemsInFilter(Boolean value) { IncludeNewItemsInFilter = value; return this; }

        public XLPivotLayout Layout { get; set; }

        public IXLPivotField SetLayout(XLPivotLayout value) { Layout = value; return this; }

        public Boolean SubtotalsAtTop { get; set; }

        public IXLPivotField SetSubtotalsAtTop() { SubtotalsAtTop = true; return this; }

        public IXLPivotField SetSubtotalsAtTop(Boolean value) { SubtotalsAtTop = value; return this; }

        public Boolean RepeatItemLabels { get; set; }

        public IXLPivotField SetRepeatItemLabels() { RepeatItemLabels = true; return this; }

        public IXLPivotField SetRepeatItemLabels(Boolean value) { RepeatItemLabels = value; return this; }

        public Boolean InsertBlankLines { get; set; }

        public IXLPivotField SetInsertBlankLines() { InsertBlankLines = true; return this; }

        public IXLPivotField SetInsertBlankLines(Boolean value) { InsertBlankLines = value; return this; }

        public Boolean ShowBlankItems { get; set; }

        public IXLPivotField SetShowBlankItems() { ShowBlankItems = true; return this; }

        public IXLPivotField SetShowBlankItems(Boolean value) { ShowBlankItems = value; return this; }

        public Boolean InsertPageBreaks { get; set; }

        public IXLPivotField SetInsertPageBreaks() { InsertPageBreaks = true; return this; }

        public IXLPivotField SetInsertPageBreaks(Boolean value) { InsertPageBreaks = value; return this; }

        public Boolean Collapsed { get; set; }

        public IXLPivotField SetCollapsed() { Collapsed = true; return this; }

        public IXLPivotField SetCollapsed(Boolean value) { Collapsed = value; return this; }

        public XLPivotSortType SortType { get; set; }

        public IXLPivotField SetSort(XLPivotSortType value) { SortType = value; return this; }

        public IList<Object> SelectedValues { get; private set; }
        public IXLPivotField AddSelectedValue(Object value)
        {
            SelectedValues.Add(value);
            return this;
        }
    }
}
