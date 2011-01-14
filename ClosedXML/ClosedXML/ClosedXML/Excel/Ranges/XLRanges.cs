using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRanges : IXLRanges, IXLStylized
    {
        private XLWorkbook workbook;
        public XLRanges(XLWorkbook workbook, IXLStyle defaultStyle)
        {
            this.workbook = workbook;
            Style = defaultStyle;
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
        public void Add(String rangeAddress)
        {
            var byExclamation = rangeAddress.Split('!');
            var wsName = byExclamation[0].Replace("'", "");
            var rng = byExclamation[1];
            var rangeToAdd = workbook.Worksheets.Worksheet(wsName).Range(rng);
            ranges.Add((XLRange)rangeToAdd);
        }
        public void Remove(IXLRange range)
        {
            ranges.RemoveAll(r => r.ToString() == range.ToString());
        }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            var retList = new List<IXLRange>();
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
                foreach (var rng in ranges)
                {
                    rng.Style = value;
                }
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

        public override string ToString()
        {
            String retVal = ranges.Aggregate(String.Empty, (agg, r)=> agg += r.ToString() + ",");
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }


    }
}
