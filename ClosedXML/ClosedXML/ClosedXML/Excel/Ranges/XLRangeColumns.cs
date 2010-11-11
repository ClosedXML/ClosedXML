using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeColumns : IXLRangeColumns
    {
        public XLRangeColumns(XLWorksheet worksheet)
        {
            Style = worksheet.Style;
        }

        List<XLRangeColumn> ranges = new List<XLRangeColumn>();

        public void Clear()
        {
            ranges.ForEach(r => r.Clear());
        }

        public void Add(IXLRangeColumn range)
        {
            ranges.Add((XLRangeColumn)range);
        }

        public IEnumerator<IXLRangeColumn> GetEnumerator()
        {
            return ranges.ToList<IXLRangeColumn>().GetEnumerator();
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

        public String FormulaA1
        {
            set
            {
                ranges.ForEach(r => r.FormulaA1 = value);
            }
        }
        public String FormulaR1C1
        {
            set
            {
                ranges.ForEach(r => r.FormulaR1C1 = value);
            }
        }
    }
}
