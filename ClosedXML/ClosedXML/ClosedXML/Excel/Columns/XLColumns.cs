using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLColumns: IXLColumns
    {
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
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public void Delete()
        {
            this.Column(ColumnNumber).Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }
    }
}
