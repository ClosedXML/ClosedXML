using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRangeRows : IXLRangeRows, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        private readonly List<XLRangeRow> _ranges = new List<XLRangeRow>();
        private IXLStyle _style;

        public XLRangeRows()
        {
            _style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        #region IXLRangeRows Members

        public IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public void Add(IXLRangeRow range)
        {
            _ranges.Add((XLRangeRow)range);
        }

        public IEnumerator<IXLRangeRow> GetEnumerator()
        {
            return _ranges.Cast<IXLRangeRow>()
                          .OrderBy(r=>r.Worksheet.Position)
                          .ThenBy(r => r.RowNumber())
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
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(true, includeFormats);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRangeRows SetDataType(XLCellValues dataType)
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
                foreach (XLRangeRow rng in _ranges)
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