using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeRows : IXLRangeRows
    {
        public XLRangeRows(XLWorksheet worksheet)
        {
            Style = worksheet.Style;
        }

        List<XLRangeRow> ranges = new List<XLRangeRow>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Add(IXLRangeRow range)
        {
            ranges.Add((XLRangeRow)range);
        }

        public IEnumerator<IXLRangeRow> GetEnumerator()
        {
            return ranges.ToList<IXLRangeRow>().GetEnumerator();
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
                        c.Address.RowNumber >= rng.FirstAddressInSheet.RowNumber
                        && c.Address.RowNumber <= rng.LastAddressInSheet.RowNumber
                        && c.Address.ColumnNumber >= rng.FirstAddressInSheet.ColumnNumber
                        && c.Address.ColumnNumber <= rng.LastAddressInSheet.ColumnNumber
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
