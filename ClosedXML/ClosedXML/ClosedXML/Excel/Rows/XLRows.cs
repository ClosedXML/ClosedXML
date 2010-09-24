using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLRows: IXLRows
    {
        public XLRows()
        {
            Style = XLWorkbook.DefaultStyle;
        }

        List<IXLRow> rows = new List<IXLRow>();
        public IEnumerator<IXLRow> GetEnumerator()
        {
            return rows.GetEnumerator();
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
                foreach (var col in rows)
                {
                    yield return col.Style;
                    foreach (var c in col.Internals.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.Row == col.Internals.FirstCellAddress.Row))
                    {
                        yield return c.Style;
                    }
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion

        public double Height
        {
            set
            {
                rows.ForEach(c => c.Height = value);
            }
        }

        public void Delete()
        {
            rows.ForEach(c => c.Delete(XLShiftDeletedCells.ShiftCellsUp));
        }


        public void Add(IXLRow row)
        {
            rows.Add(row);
        }
    }
}
