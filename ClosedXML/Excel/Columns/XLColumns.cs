using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumns :
#if !STYLES_REWORK
        XLStylizedBase,
#endif
        IXLColumns
    {
        private readonly List<XLColumn> _columnsCollection = new List<XLColumn>();

        private readonly XLWorkbook _workbook;
        private readonly XLWorksheet? _worksheet;
        private readonly XLWorksheet? _defaultStyleSheet;

        /// <summary>
        /// This object represents all columns of the worksheet, even non-materialized ones.
        /// </summary>
        [MemberNotNullWhen(true, nameof(_worksheet))]
        private bool AllColumnsOfSheet => _worksheet is not null;

        private bool IsMaterialized => _lazyEnumerable == null;

        private IEnumerable<XLColumn>? _lazyEnumerable;

        private IEnumerable<XLColumn> Columns => _lazyEnumerable ?? _columnsCollection.AsEnumerable();

        /// <summary>
        /// Create a new instance of <see cref="XLColumns"/>.
        /// </summary>
        /// <param name="workbook">Workbook to which all columns belong.</param>
        /// <param name="worksheet">If worksheet is specified it means that the created instance represents
        /// all columns on a worksheet so changing its width will affect all columns.</param>
        /// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
        /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLColumn"/> to support lazy initialization.</param>
        public XLColumns(XLWorkbook workbook, XLWorksheet? worksheet, XLWorksheet? defaultStyleSheet = null, IEnumerable<XLColumn>? lazyEnumerable = null)
#if !STYLES_REWORK
            : base(defaultStyleSheet?.StyleValue)
#endif
        {
            _workbook = workbook;
            _worksheet = worksheet;
            _defaultStyleSheet = defaultStyleSheet;
            _lazyEnumerable = lazyEnumerable;
        }

        #region IXLColumns Members

        public IEnumerator<IXLColumn> GetEnumerator()
        {
            return Columns.Cast<IXLColumn>().OrderBy(r => r.ColumnNumber()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Double Width
        {
            set
            {
                Columns.ForEach(c => c.Width = value);

                if (!AllColumnsOfSheet) return;

                _worksheet.ColumnWidth = value;
                _worksheet.Internals.ColumnsCollection.ForEach(c => c.Value.Width = value);
            }
        }

        public void Delete()
        {
            if (AllColumnsOfSheet)
            {
                _worksheet.Internals.ColumnsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (XLColumn c in Columns)
                {
                    if (!toDelete.TryGetValue(c.Worksheet, out List<Int32> list))
                    {
                        list = new List<Int32>();
                        toDelete.Add(c.Worksheet, list);
                    }

                    list.Add(c.ColumnNumber());
                }

                foreach (KeyValuePair<IXLWorksheet, List<int>> kp in toDelete)
                {
                    foreach (int c in kp.Value.OrderByDescending(c => c))
                        kp.Key.Column(c).Delete();
                }
            }
        }

        public IXLColumns AdjustToContents()
        {
            Columns.ForEach(c => c.AdjustToContents());
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow, endRow));
            return this;
        }

        public IXLColumns AdjustToContents(Double minWidth, Double maxWidth)
        {
            Columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
            return this;
        }

        public void Hide()
        {
            Columns.ForEach(c => c.Hide());
        }

        public void Unhide()
        {
            Columns.ForEach(c => c.Unhide());
        }

        public void Group()
        {
            Group(false);
        }

        public void Group(Int32 outlineLevel)
        {
            Group(outlineLevel, false);
        }

        public void Ungroup()
        {
            Ungroup(false);
        }

        public void Group(Boolean collapse)
        {
            Columns.ForEach(c => c.Group(collapse));
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            Columns.ForEach(c => c.Group(outlineLevel, collapse));
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            Columns.ForEach(c => c.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            Columns.ForEach(c => c.Collapse());
        }

        public void Expand()
        {
            Columns.ForEach(c => c.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(_workbook, false, XLCellsUsedOptions.All);
            foreach (XLColumn container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(_workbook, true, XLCellsUsedOptions.All);
            foreach (XLColumn container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(_workbook, true, options);
            foreach (XLColumn container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        /// <summary>
        ///   Adds a vertical page break after this column.
        /// </summary>
        public IXLColumns AddVerticalPageBreaks()
        {
            foreach (XLColumn col in Columns)
                col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
            return this;
        }

#if STYLES_REWORK
        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }
#endif

        internal XLCellFormat Format
        {
            get
            {
                if (AllColumnsOfSheet)
                {
                    return XLCellFormat.ForWorksheet(_worksheet);
                }

                var columns = Columns.Select(x => x.Area).ToArray();
                return XLCellFormat.ForColumns(_workbook, _defaultStyleSheet, columns);
            }
        }

        #endregion IXLColumns Members

#if !STYLES_REWORK
        #region IXLStylized Members

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                if (AllColumnsOfSheet)
                    yield return _worksheet;
                else
                {
                    foreach (XLColumn column in Columns)
                        yield return column;
                }
            }
        }

        public override IEnumerable<IXLRange> RangesUsed
        {
            get
            {
                var retVal = new XLRanges(_workbook);
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion IXLStylized Members
#endif
        public void Add(XLColumn column)
        {
            Materialize();
            _columnsCollection.Add(column);
        }

        public void CollapseOnly()
        {
            Columns.ForEach(c => c.Collapsed = true);
        }

        public IXLColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            Columns.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        private void Materialize()
        {
            if (IsMaterialized)
                return;

            _columnsCollection.AddRange(Columns);
            _lazyEnumerable = null;
        }
    }
}
