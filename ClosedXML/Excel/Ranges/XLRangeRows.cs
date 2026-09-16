#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeRows : IXLRangeRows
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLRangeRow> _ranges = new List<XLRangeRow>();

        public XLRangeRows(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        internal XLCellFormat Format
        {
            get
            {
                var areas = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
                return XLCellFormat.ForAreas(_worksheet.Workbook, areas, null);
            }
        }

        #region IXLRangeRows Members

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

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

        public IXLRangeRows AdjustToContents()
        {
            OrderedRows.ForEach(r => r.AdjustToContents());
            return this;
        }

        public IXLRangeRows AdjustToContents(Int32 startColumn)
        {
            OrderedRows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }

        public IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            OrderedRows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }

        public IXLRangeRows AdjustToContents(Double minHeight, Double maxHeight)
        {
            OrderedRows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXLRangeRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            OrderedRows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            OrderedRows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRangeRows Where(Func<IXLRangeRow, Boolean> predicate)
        {
            if (predicate is null)
                throw new ArgumentNullException(nameof(predicate));

            return FromOrdered(OrderedRows.Where(r => predicate(r)));
        }

        public IXLRangeRows Skip(Int32 count)
        {
            return FromOrdered(OrderedRows.Skip(count));
        }

        public IXLRangeRows Take(Int32 count)
        {
            return FromOrdered(OrderedRows.Take(count));
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

        private IEnumerable<XLRangeRow> OrderedRows =>
            _ranges.OrderBy(r => r.Worksheet.Position).ThenBy(r => r.RowNumber());

        private IXLRangeRows FromOrdered(IEnumerable<XLRangeRow> rows)
        {
            var result = new XLRangeRows(_worksheet);
            foreach (var row in rows)
                result.Add(row);
            return result;
        }
    }
}
