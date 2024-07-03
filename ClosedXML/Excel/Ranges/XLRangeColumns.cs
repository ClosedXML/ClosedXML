#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRangeColumns : XLStylizedBase, IXLRangeColumns, IXLStylized
    {
        private readonly List<XLRangeColumn> _ranges = new List<XLRangeColumn>();

        public XLRangeColumns() : base(XLWorkbook.DefaultStyleValue)
        {
        }

        #region IXLRangeColumns Members

        public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
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

        public IXLCells Cells()
        {
            var cells = new XLCells(usedCellsOnly: false, options: XLCellsUsedOptions.AllContents);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(usedCellsOnly: true, options: XLCellsUsedOptions.AllContents);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }


        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(usedCellsOnly: true, options: options);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        #endregion IXLRangeColumns Members

        #region IXLStylized Members
        
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
