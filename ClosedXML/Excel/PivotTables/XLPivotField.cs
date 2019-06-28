using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{SourceName}")]
    internal class XLPivotField : IXLPivotField
    {
        private readonly IXLPivotTable _pivotTable;
        public XLPivotField(IXLPivotTable pivotTable, string sourceName)
        {
            this._pivotTable = pivotTable;
            SourceName = sourceName;
            Subtotals = new List<XLSubtotalFunction>();
            SelectedValues = new List<Object>();
            SortType = XLPivotSortType.Default;
            SetExcelDefaults();

            StyleFormats = new XLPivotFieldStyleFormats(this);
        }

        public String SourceName { get; private set; }
        public String CustomName { get; set; }

        public IXLPivotField SetCustomName(String value) { CustomName = value; return this; }

        public String SubtotalCaption { get; set; }

        public IXLPivotField SetSubtotalCaption(String value) { SubtotalCaption = value; return this; }

        public List<XLSubtotalFunction> Subtotals { get; private set; }

        public IXLPivotField AddSubtotal(XLSubtotalFunction value) { Subtotals.Add(value); return this; }

        public Boolean IncludeNewItemsInFilter { get; set; }

        public IXLPivotField SetIncludeNewItemsInFilter() { IncludeNewItemsInFilter = true; return this; }

        public IXLPivotField SetIncludeNewItemsInFilter(Boolean value) { IncludeNewItemsInFilter = value; return this; }

        public bool Outline { get; set; }
        public bool Compact { get; set; }

        public IXLPivotField SetLayout(XLPivotLayout value)
        {
            Compact = false;
            Outline = false;
            switch (value)
            {
                case XLPivotLayout.Compact: Compact = true; break;
                case XLPivotLayout.Outline: Outline = true; break;
            }
            return this;
        }

        public Boolean? SubtotalsAtTop { get; set; }

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

        private void SetExcelDefaults()
        {
            IncludeNewItemsInFilter = false;
            Outline = true;
            Compact = true;
            InsertBlankLines = false;
            ShowBlankItems = false;
            InsertPageBreaks = false;
            RepeatItemLabels = false;
            SubtotalsAtTop = true;
            Collapsed = false;
        }

        public IXLPivotFieldStyleFormats StyleFormats { get; set; }

        public Boolean IsOnRowAxis => _pivotTable.RowLabels.Contains(this.SourceName);

        public Boolean IsOnColumnAxis => _pivotTable.ColumnLabels.Contains(this.SourceName);

        public Boolean IsInFilterList => _pivotTable.ReportFilters.Contains(this.SourceName);

        public Int32 Offset => _pivotTable.Source.SourceRangeFields.ToList().IndexOf(this.SourceName);
    }
}
