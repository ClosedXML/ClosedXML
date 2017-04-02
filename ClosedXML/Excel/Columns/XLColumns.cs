using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLColumns : IXLColumns, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        private readonly List<XLColumn> _columns = new List<XLColumn>();
        private readonly XLWorksheet _worksheet;
        internal IXLStyle style;

        public XLColumns(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            style = new XLStyle(this, XLWorkbook.DefaultStyle);
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

        public IXLStyle Style
        {
            get { return style; }
            set
            {
                style = new XLStyle(this, value);

                if (_worksheet != null)
                    _worksheet.Style = value;
                else
                {
                    foreach (XLColumn column in _columns)
                        column.Style = value;
                }
            }
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
            var cells = new XLCells(false, false);
            foreach (XLColumn container in _columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLColumn container in _columns)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(true, includeFormats);
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

        public IXLColumns SetDataType(XLCellValues dataType)
        {
            _columns.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                if (_worksheet != null)
                    yield return _worksheet.Style;
                else
                {
                    foreach (IXLStyle s in _columns.SelectMany(col => col.Styles))
                    {
                        yield return s;
                    }
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return style; }
            set { style = new XLStyle(this, value); }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion

        public void Add(XLColumn column)
        {
            _columns.Add(column);
        }

        public void CollapseOnly()
        {
            _columns.ForEach(c => c.Collapsed = true);
        }

        public IXLColumns Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _columns.ForEach(c=>c.Clear(clearOptions));
            return this;
        }

        public void Dispose()
        {
            if (_columns != null)
                _columns.ForEach(c => c.Dispose());
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}