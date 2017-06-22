using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRows : IXLRows, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        private readonly List<XLRow> _rows = new List<XLRow>();
        private readonly XLWorksheet _worksheet;
        internal IXLStyle style;

        public XLRows(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        #region IXLRows Members

        public IEnumerator<IXLRow> GetEnumerator()
        {
            return _rows.Cast<IXLRow>().OrderBy(r=>r.RowNumber()).GetEnumerator();
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
                    foreach (XLRow row in _rows)
                        row.Style = value;
                }
            }
        }

        public double Height
        {
            set
            {
                _rows.ForEach(c => c.Height = value);
                if (_worksheet == null) return;
                _worksheet.RowHeight = value;
                _worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
            }
        }

        public void Delete()
        {
            if (_worksheet != null)
            {
                _worksheet.Internals.RowsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (XLRow r in _rows)
                {
                    if (!toDelete.ContainsKey(r.Worksheet))
                        toDelete.Add(r.Worksheet, new List<Int32>());

                    toDelete[r.Worksheet].Add(r.RowNumber());
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
            _rows.ForEach(r => r.AdjustToContents());
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn)
        {
            _rows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            _rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }

        public IXLRows AdjustToContents(Double minHeight, Double maxHeight)
        {
            _rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            _rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            _rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
            return this;
        }


        public void Hide()
        {
            _rows.ForEach(r => r.Hide());
        }

        public void Unhide()
        {
            _rows.ForEach(r => r.Unhide());
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
            _rows.ForEach(r => r.Group(collapse));
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            _rows.ForEach(r => r.Group(outlineLevel, collapse));
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            _rows.ForEach(r => r.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            _rows.ForEach(r => r.Collapse());
        }

        public void Expand()
        {
            _rows.ForEach(r => r.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, false);
            foreach (XLRow container in _rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLRow container in _rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(true, includeFormats);
            foreach (XLRow container in _rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRows AddHorizontalPageBreaks()
        {
            foreach (XLRow row in _rows)
                row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
            return this;
        }

        public IXLRows SetDataType(XLCellValues dataType)
        {
            _rows.ForEach(c => c.DataType = dataType);
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
                    foreach (IXLStyle s in _rows.SelectMany(row => row.Styles))
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

        public void Add(XLRow row)
        {
            _rows.Add(row);
        }

        public IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _rows.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Dispose()
        {
            if (_rows != null)
                _rows.ForEach(r => r.Dispose());
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}