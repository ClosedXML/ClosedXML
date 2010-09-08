using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public class XLRow: IXLRow
    {
        public XLRow(Int32 row, XLRowParameters xlRowParameters)
        {
            Internals = new XLRangeInternals(new XLAddress(row, 1), new XLAddress(row, XLWorksheet.MaxNumberOfColumns), xlRowParameters.Worksheet);
            RowNumber = row;
            ColumnNumber = 1;
            ColumnLetter = "A";
            this.style = new XLStyle(this, xlRowParameters.DefaultStyle);
            this.Height = XLWorkbook.DefaultRowHeight;
        }

        public Double Height { get; set; }
        public Int32 RowNumber { get; private set; }
        public Int32 ColumnNumber { get; private set; }
        public String ColumnLetter { get; private set; }

        #region IXLRange Members

        public IXLRangeInternals Internals { get; private set; }

        #endregion

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
                foreach (var c in Internals.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.Row == Internals.FirstCellAddress.Row))
                {
                    yield return c.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion

        #region IXLRange Members

        public IXLRange Row(int row)
        {
            return this;
        }

        public IXLRange Column(int column)
        {
            var address = new XLAddress(1, column);
            return this.Range(address, address);
        }

        public IXLRange Column(string column)
        {
            return Column(Int32.Parse(column));
        }

        #endregion
    }
}
