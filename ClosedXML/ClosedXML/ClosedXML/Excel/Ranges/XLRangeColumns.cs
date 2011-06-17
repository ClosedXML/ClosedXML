using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumns : IXLRangeColumns, IXLStylized
    {
        public XLRangeColumns()
        {
            style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        List<XLRangeColumn> ranges = new List<XLRangeColumn>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Delete()
        {
            ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
            ranges.Clear();
        }

        public void Add(IXLRangeColumn range)
        {
            ranges.Add((XLRangeColumn)range);
        }

        public IEnumerator<IXLRangeColumn> GetEnumerator()
        {
            var retList = new List<IXLRangeColumn>();
            ranges.ForEach(c => retList.Add(c));
            return retList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #region IXLStylized Members

        private IXLStyle style;
        public IXLStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = new XLStyle(this, value);
                ranges.ForEach(r => r.Style = value);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                foreach (var rng in ranges)
                {
                    yield return rng.Style;
                    foreach (var r in (rng.Worksheet as XLWorksheet).Internals.CellsCollection.Values.Where(c =>
                        c.Address.RowNumber >= rng.RangeAddress.FirstAddress.RowNumber
                        && c.Address.RowNumber <= rng.RangeAddress.LastAddress.RowNumber
                        && c.Address.ColumnNumber >= rng.RangeAddress.FirstAddress.ColumnNumber
                        && c.Address.ColumnNumber <= rng.RangeAddress.LastAddress.ColumnNumber
                        ))
                    {
                        yield return r.Style;
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

        #endregion


        public IXLCells Cells()
        {
            var cells = new XLCells(false, false, false);
            foreach (var container in ranges)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(false, true, false);
            foreach (var container in ranges)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(false, true, includeStyles);
            foreach (var container in ranges)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
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

        public IXLRangeColumns SetDataType(XLCellValues dataType)
        {
            ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

    }
}
