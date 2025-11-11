using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRows :
#if !STYLES_REWORK
        XLStylizedBase,
#endif
        IXLRows
    {
        private readonly List<XLRow> _rowsCollection = new List<XLRow>();
        private readonly XLWorkbook _workbook;
        private readonly XLWorksheet? _worksheet;
        private readonly XLWorksheet? _defaultStyleSheet;

        /// <summary>
        /// This object represents all rows of the worksheet, even non-materialized ones.
        /// </summary>
        [MemberNotNullWhen(true, nameof(_worksheet))]
        private bool AllRowsOfSheet => _worksheet is not null;

        private bool IsMaterialized => _lazyEnumerable == null;

        private IEnumerable<XLRow>? _lazyEnumerable;
        private IEnumerable<XLRow> Rows => _lazyEnumerable ?? _rowsCollection.AsEnumerable();

        /// <summary>
        /// Create a new instance of <see cref="XLRows"/>.
        /// </summary>
        /// <param name="workbook">Workbook of the rows.</param>
        /// <param name="worksheet">If worksheet is specified it means that the created instance represents
        /// all rows on a worksheet so changing its height will affect all rows.</param>
        /// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
        /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLRow"/> to support lazy initialization.</param>
        public XLRows(XLWorkbook workbook, XLWorksheet? worksheet, XLWorksheet? defaultStyleSheet = null, IEnumerable<XLRow>? lazyEnumerable = null)
#if !STYLES_REWORK
            : base(defaultStyleSheet?.StyleValue)
#endif
        {
            _workbook = workbook;
            _worksheet = worksheet;
            _defaultStyleSheet = defaultStyleSheet;
            _lazyEnumerable = lazyEnumerable;
        }

        #region IXLRows Members

        public IEnumerator<IXLRow> GetEnumerator()
        {
            return Rows.Cast<IXLRow>().OrderBy(r => r.RowNumber()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public double Height
        {
            set
            {
                Rows.ForEach(c => c.Height = value);
                if (!AllRowsOfSheet)
                    return;

                _worksheet.RowHeight = value;
                _worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
            }
        }

        public void Delete()
        {
            if (AllRowsOfSheet)
            {
                _worksheet.Internals.RowsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (XLRow r in Rows)
                {
                    if (!toDelete.TryGetValue(r.Worksheet, out List<Int32> list))
                    {
                        list = new List<Int32>();
                        toDelete.Add(r.Worksheet, list);
                    }

                    list.Add(r.RowNumber());
                }

                foreach (KeyValuePair<IXLWorksheet, List<int>> kp in toDelete)
                {
                    foreach (int r in kp.Value.OrderByDescending(r => r))
                        kp.Key.Row(r).Delete();
                }
            }
        }

        public IXLRows AdjustToContents()
        {
            Rows.ForEach(r => r.AdjustToContents());
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }

        public IXLRows AdjustToContents(Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
            return this;
        }

        public void Hide()
        {
            Rows.ForEach(r => r.Hide());
        }

        public void Unhide()
        {
            Rows.ForEach(r => r.Unhide());
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
            Rows.ForEach(r => r.Group(collapse));
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            Rows.ForEach(r => r.Group(outlineLevel, collapse));
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            Rows.ForEach(r => r.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            Rows.ForEach(r => r.Collapse());
        }

        public void Expand()
        {
            Rows.ForEach(r => r.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(_workbook, false, XLCellsUsedOptions.AllContents);
            foreach (XLRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(_workbook, true, XLCellsUsedOptions.AllContents);
            foreach (XLRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(_workbook, true, options);
            foreach (XLRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRows AddHorizontalPageBreaks()
        {
            foreach (XLRow row in Rows)
                row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
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
                if (AllRowsOfSheet)
                {
                    return XLCellFormat.ForWorksheet(_worksheet);
                }

                return XLCellFormat.ForRows(_workbook, _defaultStyleSheet, Rows);

            }
        }

        #endregion IXLRows Members

#if !STYLES_REWORK
        #region IXLStylized Members

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                if (AllRowsOfSheet)
                    yield return _worksheet;
                else
                {
                    foreach (XLRow row in Rows)
                        yield return row;
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

        public void Add(XLRow row)
        {
            Materialize();
            _rowsCollection.Add(row);
        }

        public IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            Rows.ForEach(c => c.Clear(clearOptions));
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

            _rowsCollection.AddRange(Rows);
            _lazyEnumerable = null;
        }
    }
}
