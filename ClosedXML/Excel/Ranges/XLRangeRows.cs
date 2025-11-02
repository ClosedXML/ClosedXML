#nullable disable

using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeRows :
#if !STYLES_REWORK
        XLStylizedBase,
#endif
        IXLRangeRows
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLRangeRow> _ranges = new List<XLRangeRow>();

        public XLRangeRows(XLWorksheet worksheet)
#if !STYLES_REWORK
            : base(XLStyle.Default.Value)
#endif
        {
            _worksheet = worksheet;
        }

        internal XLCellFormat Format
        {
            get
            {
                var areas = _ranges.Select(x => XLBookArea.From(x.RangeAddress)).ToArray();
                return XLCellFormat.ForAreas(_worksheet.Workbook, areas, null);
            }
        }

        #region IXLRangeRows Members

#if STYLES_REWORK
        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }
#endif

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
            var cells = new XLCells(_worksheet, false, XLCellsUsedOptions.AllContents);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(_worksheet, true, XLCellsUsedOptions.AllContents);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }


        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(_worksheet, true, options);
            foreach (XLRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        #endregion IXLRangeRows Members

#if !STYLES_REWORK
        #region IXLStylized Members

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (var range in _ranges)
                    yield return range;
            }
        }

        public override IEnumerable<IXLRange> RangesUsed
        {
            get
            {
                var retVal = new XLRanges(_worksheet);
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion IXLStylized Members
#endif
    }
}
