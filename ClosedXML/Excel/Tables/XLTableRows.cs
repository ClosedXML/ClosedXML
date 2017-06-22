using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLTableRows : IXLTableRows, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        private readonly List<XLTableRow> _ranges = new List<XLTableRow>();
        private IXLStyle _style;
        

        public XLTableRows(IXLStyle defaultStyle)
        {
            _style = new XLStyle(this, defaultStyle);
        }

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return _style;
                foreach (XLTableRow rng in _ranges)
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

        #region IXLTableRows Members

        public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _ranges.ForEach(r => r.Clear(clearOptions));
            return this;
        }

        public void Add(IXLTableRow range)
        {
            _ranges.Add((XLTableRow)range);
        }

        public IEnumerator<IXLTableRow> GetEnumerator()
        {
            var retList = new List<IXLTableRow>();
            _ranges.ForEach(retList.Add);
            return retList.GetEnumerator();
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
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(false, includeFormats);
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        #endregion

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}