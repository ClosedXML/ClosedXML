using ClosedXML.Excel.Style;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLTableRows : XLStylizedBase, IXLTableRows, IXLStylized
    {
        private readonly List<XLTableRow> _ranges = new List<XLTableRow>();

        public XLTableRows(IXLStyle defaultStyle) : base((defaultStyle as XLStyle).Value)
        {
        }

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
                foreach (var rng in _ranges)
                {
                    yield return rng.Style;
                    foreach (var r in rng.Worksheet.Internals.CellsCollection.GetCells(
                        rng.RangeAddress.FirstAddress.RowNumber,
                        rng.RangeAddress.FirstAddress.ColumnNumber,
                        rng.RangeAddress.LastAddress.RowNumber,
                        rng.RangeAddress.LastAddress.ColumnNumber))
                    {
                        yield return r.Style;
                    }
                }
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (var range in _ranges)
                {
                    yield return range;
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

        #region IXLTableRows Members

        public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _ranges.ForEach(r => r.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public void Add(IXLTableRow tableRow)
        {
            _ranges.Add((XLTableRow)tableRow);
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

        public IXLCells Cells()
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (var container in _ranges)
            {
                cells.Add(container.RangeAddress);
            }

            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents);
            foreach (var container in _ranges)
            {
                cells.Add(container.RangeAddress);
            }

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
            var cells = new XLCells(false, options);
            foreach (var container in _ranges)
            {
                cells.Add(container.RangeAddress);
            }

            return cells;
        }

        #endregion IXLTableRows Members

        public void Select()
        {
            foreach (var range in this)
            {
                range.Select();
            }
        }
    }
}
