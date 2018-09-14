using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRangeRows : XLStylizedBase, IXLRangeRows, IXLStylized
    {
        private readonly List<XLRangeRow> _ranges = new List<XLRangeRow>();

        public XLRangeRows() : base(XLStyle.Default.Value)
        {
        }

        #region IXLRangeRows Members

        public IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
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
                          .OrderBy(r => r.Worksheet.Position)
                          .ThenBy(r => r.RowNumber())
                          .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(true, options);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRangeRows SetDataType(XLDataType dataType)
        {
            _ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion IXLRangeRows Members

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
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
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (var range in _ranges)
                    yield return range;
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

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}
