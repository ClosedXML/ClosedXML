#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLAutoFilter : IXLAutoFilter
    {
        /// <summary>
        /// Key is column number.
        /// </summary>
        private readonly Dictionary<Int32, XLFilterColumn> _columns = new();

        internal IReadOnlyDictionary<Int32, XLFilterColumn> Columns => _columns;

        #region IXLAutoFilter Members

        public IEnumerable<IXLRangeRow> HiddenRows => Range.Rows(r => r.WorksheetRow().IsHidden);

        public Boolean IsEnabled { get; set; }

        public IXLRange Range { get; set; }

        public Int32 SortColumn { get; set; }

        public Boolean Sorted { get; set; }

        public XLSortOrder SortOrder { get; set; }

        public IEnumerable<IXLRangeRow> VisibleRows => Range.Rows(r => !r.WorksheetRow().IsHidden);

        IXLAutoFilter IXLAutoFilter.Clear() => Clear();

        IXLFilterColumn IXLAutoFilter.Column(String columnLetter) => Column(columnLetter);

        IXLFilterColumn IXLAutoFilter.Column(Int32 columnNumber) => Column(columnNumber);

        IXLAutoFilter IXLAutoFilter.Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public IXLAutoFilter Reapply()
        {
            // Recalculate shown / hidden rows
            var rows = Range.Rows(2, Range.RowCount());
            rows.ForEach(row =>
                row.WorksheetRow().Unhide()
            );

            foreach (var filterColumn in _columns.Values)
                filterColumn.Refresh();

            foreach (IXLRangeRow row in rows)
            {
                var rowMatch = true;

                foreach (var (columnIndex, column) in _columns)
                {
                    var cell = row.Cell(columnIndex);
                    var columnFilterMatch = column.Check(cell);
                    rowMatch &= columnFilterMatch;

                    if (!rowMatch) break;
                }

                if (!rowMatch) row.WorksheetRow().Hide();
            }

            return this;
        }

        #endregion IXLAutoFilter Members

        internal XLFilterColumn Column(String columnLetter)
        {
            var columnNumber = XLHelper.GetColumnNumberFromLetter(columnLetter);
            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(columnLetter), "Column '" + columnLetter + "' is outside the allowed column range.");

            return Column(columnNumber);
        }

        internal XLFilterColumn Column(Int32 columnNumber)
        {
            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column " + columnNumber + " is outside the allowed column range.");

            if (!_columns.TryGetValue(columnNumber, out XLFilterColumn filterColumn))
            {
                filterColumn = new XLFilterColumn(this, columnNumber);
                _columns.Add(columnNumber, filterColumn);
            }

            return filterColumn;
        }

        internal XLAutoFilter Clear()
        {
            if (!IsEnabled) return this;

            IsEnabled = false;
            foreach (var filterColumn in _columns.Values)
                filterColumn.Clear(false);

            foreach (IXLRangeRow row in Range.Rows().Where(r => r.RowNumber() > 1))
                row.WorksheetRow().Unhide();
            return this;
        }

        internal XLAutoFilter Set(IXLRangeBase range)
        {
            var firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t => t.RangeUsed().Intersects(range));
            if (firstOverlappingTable != null)
                throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

            Range = range.AsRange();
            IsEnabled = true;
            return this;
        }

        internal XLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder, Boolean matchCase, Boolean ignoreBlanks)
        {
            if (!IsEnabled)
                throw new InvalidOperationException("Filter has not been enabled.");

            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
                                                                              ignoreBlanks);

            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            Reapply();

            return this;
        }
    }
}
