using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRanges : IXLRanges
    {
        public XLRanges(XLWorksheet worksheet)
        {
            Style = worksheet.Style;
        }

        List<XLRange> ranges = new List<XLRange>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Add(IXLRange range)
        {
            ranges.Add((XLRange)range);
        }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            return ranges.ToList<IXLRange>().GetEnumerator();
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

        #endregion
    }
}
