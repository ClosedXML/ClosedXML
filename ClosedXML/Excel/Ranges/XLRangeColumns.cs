using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRangeColumns : IXLRangeColumns, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        private readonly List<XLRangeColumn> _ranges = new List<XLRangeColumn>();
        private IXLStyle _style;

        public XLRangeColumns()
        {
            _style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        #region IXLRangeColumns Members

        public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            _ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public void Add(IXLRangeColumn range)
        {
            _ranges.Add((XLRangeColumn)range);
        }

        public IEnumerator<IXLRangeColumn> GetEnumerator()
        {
            return _ranges.Cast<IXLRangeColumn>()
              .OrderBy(r => r.Worksheet.Position)
              .ThenBy(r => r.ColumnNumber())
              .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLStyle Style
        {
            get { return _style; }
            set
            {
                _style = new XLStyle(this, value);
                _ranges.ForEach(r => r.Style = value);
            }
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, false);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(true, includeFormats);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRangeColumns SetDataType(XLCellValues dataType)
        {
            _ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return _style;
                foreach (XLRangeColumn rng in _ranges)
                {
                    yield return rng.Style;
                    foreach (XLCell r in rng.Worksheet.Internals.CellsCollection.GetCells(
                        rng.RangeAddress.FirstAddress.RowNumber,
                        rng.RangeAddress.FirstAddress.ColumnNumber,
                        rng.RangeAddress.LastAddress.RowNumber,
                        rng.RangeAddress.LastAddress.ColumnNumber))
                        yield return r.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return _style; }
            set { _style = new XLStyle(this, value); }
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

        public void Dispose()
        {
            if (_ranges != null)
                _ranges.ForEach(r => r.Dispose());
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}