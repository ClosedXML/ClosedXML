using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumns : XLStylizedBase, IXLColumns, IXLStylized
    {
        private readonly List<XLColumn> _columnsCollection = new List<XLColumn>();
        private readonly XLWorksheet _worksheet;
        private bool IsMaterialized => _lazyEnumerable == null;

        private IEnumerable<XLColumn> _lazyEnumerable;
        private IEnumerable<XLColumn> Columns => _lazyEnumerable ?? _columnsCollection.AsEnumerable();

        /// <summary>
        /// Create a new instance of <see cref="XLColumns"/>.
        /// </summary>
        /// <param name="worksheet">If worksheet is specified it means that the created instance represents
        /// all columns on a worksheet so changing its width will affect all columns.</param>
        /// <param name="defaultStyle">Default style to use when initializing child entries.</param>
        /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLColumn"/> to support lazy initialization.</param>
        public XLColumns(XLWorksheet worksheet, XLStyleValue defaultStyle = null, IEnumerable<XLColumn> lazyEnumerable = null)
            : base(defaultStyle)
        {
            _worksheet = worksheet;
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

        public double Width
        {
            set
            {
                Columns.ForEach(c => c.Width = value);

                if (_worksheet == null) return;

                _worksheet.ColumnWidth = value;
                _worksheet.Internals.ColumnsCollection.ForEach(c => c.Value.Width = value);
            }
        }

        public void Delete()
        {
            if (_worksheet != null)
            {
                _worksheet.Internals.ColumnsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<int>>();
                foreach (var c in Columns)
                {
                    if (!toDelete.TryGetValue(c.Worksheet, out var list))
                    {
                        list = new List<int>();
                        toDelete.Add(c.Worksheet, list);
                    }

                    list.Add(c.ColumnNumber());
                }

                foreach (var kp in toDelete)
                {
                    foreach (var c in kp.Value.OrderByDescending(c => c))
                        kp.Key.Column(c).Delete();
                }
            }
        }

        public IXLColumns AdjustToContents()
        {
            Columns.ForEach(c => c.AdjustToContents());
            return this;
        }

        public IXLColumns AdjustToContents(int startRow)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow));
            return this;
        }

        public IXLColumns AdjustToContents(int startRow, int endRow)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow, endRow));
            return this;
        }

        public IXLColumns AdjustToContents(double minWidth, double maxWidth)
        {
            Columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(int startRow, double minWidth, double maxWidth)
        {
            Columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth)
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

        public void Group(int outlineLevel)
        {
            Group(outlineLevel, false);
        }

        public void Ungroup()
        {
            Ungroup(false);
        }

        public void Group(bool collapse)
        {
            Columns.ForEach(c => c.Group(collapse));
        }

        public void Group(int outlineLevel, bool collapse)
        {
            Columns.ForEach(c => c.Group(outlineLevel, collapse));
        }

        public void Ungroup(bool ungroupFromAll)
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
            var cells = new XLCells(false, XLCellsUsedOptions.All);
            foreach (var container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.All);
            foreach (var container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(bool includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(true, options);
            foreach (var container in Columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        /// <summary>
        ///   Adds a vertical page break after this column.
        /// </summary>
        public IXLColumns AddVerticalPageBreaks()
        {
            foreach (var col in Columns)
                col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
            return this;
        }

        public IXLColumns SetDataType(XLDataType dataType)
        {
            Columns.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion IXLColumns Members

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
                if (_worksheet != null)
                    yield return _worksheet.Style;
                else
                {
                    foreach (var s in Columns.SelectMany(col => col.Styles))
                    {
                        yield return s;
                    }
                }
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                if (_worksheet != null)
                    yield return _worksheet;
                else
                {
                    foreach (var column in Columns)
                        yield return column;
                }
            }
        }

        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion IXLStylized Members

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