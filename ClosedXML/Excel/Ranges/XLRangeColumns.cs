// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRangeColumns : XLStylizedBase, IXLRangeColumns, IXLStylized
    {
        private readonly List<XLRangeColumn> _ranges = new List<XLRangeColumn>();
        private IXLRangeColumn _firstColumn;
        private IXLRangeColumn _lastColumn;

        public XLRangeColumns() : base(XLWorkbook.DefaultStyleValue)
        {
        }

        #region IXLRangeColumns Members

        public void Add(IXLRangeColumn rangeColumn)
        {
            _ranges.Add((XLRangeColumn)rangeColumn);

            if (rangeColumn.ColumnNumber() < (_firstColumn?.ColumnNumber() ?? int.MaxValue))
                _firstColumn = rangeColumn;

            if (rangeColumn.ColumnNumber() > (_lastColumn?.ColumnNumber() ?? 0))
                _lastColumn = rangeColumn;
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

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents
            );
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(usedCellsOnly: true, options: options);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public IXLRangeColumns Contiguous()
        {
            var rangeColumns = new XLRangeColumns();

            if (_ranges.Count > 0)
            {
                var ws = (XLWorksheet)_firstColumn.Worksheet;

                int previousColumnNumber = _firstColumn.ColumnNumber();

                foreach (var c in this)
                {
                    while (previousColumnNumber < c.ColumnNumber() - 1)
                    {
                        previousColumnNumber++;
                        var firstCellAddress = new XLAddress(ws, _firstColumn.RangeAddress.FirstAddress.RowNumber, previousColumnNumber, fixedRow: false, fixedColumn: false);
                        var lastCellAddress = new XLAddress(ws, _firstColumn.RangeAddress.LastAddress.RowNumber, previousColumnNumber, fixedRow: false, fixedColumn: false);
                        var column = ws.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
                        rangeColumns.Add(column);
                    }
                    rangeColumns.Add(c);
                    previousColumnNumber = c.ColumnNumber();
                }
            }

            return rangeColumns;
        }

        public void Delete()
        {
            _ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
            _firstColumn = null;
            _lastColumn = null;
        }

        public IXLRangeColumn FirstColumn() => _firstColumn;

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

        public IXLRangeColumn LastColumn() => _lastColumn;

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        public IXLRangeColumns SetDataType(XLDataType dataType)
        {
            _ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion IXLRangeColumns Members

        #region IXLStylized Members

        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
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

        #endregion IXLStylized Members
    }
}
