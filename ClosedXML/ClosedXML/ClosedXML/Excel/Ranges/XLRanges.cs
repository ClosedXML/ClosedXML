using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLRanges: IXLRanges
    {
        public XLRanges()
        {
            Style = XLWorkbook.DefaultStyle;
        }

        List<IXLRange> ranges = new List<IXLRange>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Add(IXLRange range)
        {
            ranges.Add(range);
        }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            return ranges.GetEnumerator();
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
                    foreach (var r in rng.Internals.Worksheet.Internals.CellsCollection.Values.Where(c =>
                        c.Address.Row >= rng.Internals.FirstCellAddress.Row
                        && c.Address.Row <= rng.Internals.LastCellAddress.Row
                        && c.Address.Column >= rng.Internals.FirstCellAddress.Column
                        && c.Address.Column <= rng.Internals.LastCellAddress.Column
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
