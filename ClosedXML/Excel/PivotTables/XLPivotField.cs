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
            _pivotTable = pivotTable;
            SourceName = sourceName;
            Subtotals = new List<XLSubtotalFunction>();
            SelectedValues = new List<object>();
            SortType = XLPivotSortType.Default;
            SetExcelDefaults();

            StyleFormats = new XLPivotFieldStyleFormats(this);
        }

        public string SourceName { get; private set; }
        public string CustomName { get; set; }

        public IXLPivotField SetCustomName(string value) { CustomName = value; return this; }

        public string SubtotalCaption { get; set; }

        public IXLPivotField SetSubtotalCaption(string value) { SubtotalCaption = value; return this; }

        public List<XLSubtotalFunction> Subtotals { get; private set; }

        public IXLPivotField AddSubtotal(XLSubtotalFunction value) { Subtotals.Add(value); return this; }

        public bool IncludeNewItemsInFilter { get; set; }

        public IXLPivotField SetIncludeNewItemsInFilter() { IncludeNewItemsInFilter = true; return this; }

        public IXLPivotField SetIncludeNewItemsInFilter(bool value) { IncludeNewItemsInFilter = value; return this; }

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

        public bool? SubtotalsAtTop { get; set; }

        public IXLPivotField SetSubtotalsAtTop() { SubtotalsAtTop = true; return this; }

        public IXLPivotField SetSubtotalsAtTop(bool value) { SubtotalsAtTop = value; return this; }

        public bool RepeatItemLabels { get; set; }

        public IXLPivotField SetRepeatItemLabels() { RepeatItemLabels = true; return this; }

        public IXLPivotField SetRepeatItemLabels(bool value) { RepeatItemLabels = value; return this; }

        public bool InsertBlankLines { get; set; }

        public IXLPivotField SetInsertBlankLines() { InsertBlankLines = true; return this; }

        public IXLPivotField SetInsertBlankLines(bool value) { InsertBlankLines = value; return this; }

        public bool ShowBlankItems { get; set; }

        public IXLPivotField SetShowBlankItems() { ShowBlankItems = true; return this; }

        public IXLPivotField SetShowBlankItems(bool value) { ShowBlankItems = value; return this; }

        public bool InsertPageBreaks { get; set; }

        public IXLPivotField SetInsertPageBreaks() { InsertPageBreaks = true; return this; }

        public IXLPivotField SetInsertPageBreaks(bool value) { InsertPageBreaks = value; return this; }

        public bool Collapsed { get; set; }

        public IXLPivotField SetCollapsed() { Collapsed = true; return this; }

        public IXLPivotField SetCollapsed(bool value) { Collapsed = value; return this; }

        public XLPivotSortType SortType { get; set; }

        public IXLPivotField SetSort(XLPivotSortType value) { SortType = value; return this; }

        public IList<object> SelectedValues { get; private set; }

        public IXLPivotField AddSelectedValue(object value)
        {
            SelectedValues.Add(value);
            return this;
        }

        public IXLPivotField AddSelectedValues(IEnumerable<object> values)
        {
            ((List<object>)SelectedValues).AddRange(values);
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

        public bool IsOnRowAxis => _pivotTable.RowLabels.Contains(SourceName);

        public bool IsOnColumnAxis => _pivotTable.ColumnLabels.Contains(SourceName);

        public bool IsInFilterList => _pivotTable.ReportFilters.Contains(SourceName);

        public int Offset => _pivotTable.SourceRangeFieldsAvailable.ToList().IndexOf(SourceName);
    }
}
