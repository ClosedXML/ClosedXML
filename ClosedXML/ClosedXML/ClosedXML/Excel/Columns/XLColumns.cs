using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLColumns: IXLColumns
    {
        public XLColumns()
        {
            Style = XLWorkbook.DefaultStyle;
        }

        List<IXLColumn> columns = new List<IXLColumn>();
        public IEnumerator<IXLColumn> GetEnumerator()
        {
            return columns.GetEnumerator();
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
                foreach (var col in columns)
                {
                    yield return col.Style;
                    foreach (var c in col.Internals.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.Column == col.Internals.FirstCellAddress.Column))
                    {
                        yield return c.Style;
                    }
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion

        public double Width
        {
            set
            {
                columns.ForEach(c => c.Width = value);
            }
        }

        public void Delete()
        {
            columns.ForEach(c => c.Delete(XLShiftDeletedCells.ShiftCellsLeft));
        }


        public void Add(IXLColumn column)
        {
            columns.Add(column);
        }
    }
}
