using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLColumns : XLStylizedBase, IXLColumns, IXLStylized
    {
        private readonly List<XLColumn> _columns = new List<XLColumn>();
        private readonly XLWorksheet _worksheet;

        public XLColumns(XLWorksheet worksheet)
            : base(XLStyle.Default.Value)
        {
            _worksheet = worksheet;
        }

        #region IXLColumns Members

        public IEnumerator<IXLColumn> GetEnumerator()
        {
            return _columns.Cast<IXLColumn>().OrderBy(r => r.ColumnNumber()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Double Width
        {
            set
            {
                _columns.ForEach(c => c.Width = value);

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
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (XLColumn c in _columns)
                {
                    if (!toDelete.ContainsKey(c.Worksheet))
                        toDelete.Add(c.Worksheet, new List<Int32>());

                    toDelete[c.Worksheet].Add(c.ColumnNumber());
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
            _columns.ForEach(c => c.AdjustToContents());
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow)
        {
            _columns.ForEach(c => c.AdjustToContents(startRow));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow)
        {
            _columns.ForEach(c => c.AdjustToContents(startRow, endRow));
            return this;
        }

        public IXLColumns AdjustToContents(Double minWidth, Double maxWidth)
        {
            _columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            _columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
            return this;
        }

        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            _columns.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
            return this;
        }

        public void Hide()
        {
            _columns.ForEach(c => c.Hide());
        }

        public void Unhide()
        {
            _columns.ForEach(c => c.Unhide());
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
            _columns.ForEach(c => c.Group(collapse));
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            _columns.ForEach(c => c.Group(outlineLevel, collapse));
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            _columns.ForEach(c => c.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            _columns.ForEach(c => c.Collapse());
        }

        public void Expand()
        {
            _columns.ForEach(c => c.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, XLCellsUsedOptions.All);
            foreach (XLColumn container in _columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.All);
            foreach (XLColumn container in _columns)
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
            var cells = new XLCells(true, options);
            foreach (XLColumn container in _columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        /// <summary>
        ///   Adds a vertical page break after this column.
        /// </summary>
        public IXLColumns AddVerticalPageBreaks()
        {
            foreach (XLColumn col in _columns)
                col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
            return this;
        }

        public IXLColumns SetDataType(XLDataType dataType)
        {
            _columns.ForEach(c => c.DataType = dataType);
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
                    foreach (IXLStyle s in _columns.SelectMany(col => col.Styles))
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
                    foreach (XLColumn column in _columns)
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
            _columns.Add(column);
        }

        public void CollapseOnly()
        {
            _columns.ForEach(c => c.Collapsed = true);
        }

        public IXLColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _columns.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}
