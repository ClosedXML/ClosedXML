using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeRows : IXLRangeRows, IXLStylized
    {
        XLWorksheet worksheet;
        public XLRangeRows(XLWorksheet worksheet)
        {
            style = new XLStyle(this, worksheet.Style);
            this.worksheet = worksheet;
        }

        List<XLRangeRow> ranges = new List<XLRangeRow>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Delete()
        {
            ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
            ranges.Clear();
        }

        public void Add(IXLRangeRow range)
        {
            ranges.Add((XLRangeRow)range);
        }

        public IEnumerator<IXLRangeRow> GetEnumerator()
        {
            var retList = new List<IXLRangeRow>();
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
                    foreach (var r in rng.Worksheet.Internals.CellsCollection.Values.Where(c =>
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
            var cells = new XLCells(worksheet, false, false, false);
            foreach (var container in ranges)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(worksheet, false, true, false);
            foreach (var container in ranges)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(worksheet, false, true, includeStyles);
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
                var retVal = new XLRanges(worksheet.Internals.Workbook, this.Style);
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        public IXLRangeRows Replace(String oldValue, String newValue)
        {
            ranges.ForEach(r => r.Replace(oldValue, newValue));
            return this;
        }
        public IXLRangeRows Replace(String oldValue, String newValue, XLSearchContents searchContents)
        {
            ranges.ForEach(r => r.Replace(oldValue, newValue, searchContents));
            return this;
        }
        public IXLRangeRows Replace(String oldValue, String newValue, XLSearchContents searchContents, Boolean useRegularExpressions)
        {
            ranges.ForEach(r => r.Replace(oldValue, newValue, searchContents, useRegularExpressions));
            return this;
        }
    }
}
