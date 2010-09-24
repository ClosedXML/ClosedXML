using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public class XLColumn: IXLColumn
    {
        public XLColumn(Int32 column, XLColumnParameters xlColumnParameters)
        {
            Internals = new XLRangeInternals(new XLAddress(1, column), new XLAddress(XLWorksheet.MaxNumberOfRows, column), xlColumnParameters.Worksheet);
            RowNumber = 1;
            ColumnNumber = column;
            ColumnLetter = XLAddress.GetColumnLetterFromNumber(column);
            this.style = new XLStyle(this, xlColumnParameters.DefaultStyle);
            this.Width = XLWorkbook.DefaultColumnWidth;
        }

        public Double Width { get; set; }
        public Int32 RowNumber { get; private set; }
        public Int32 ColumnNumber { get; private set; }
        public String ColumnLetter { get; private set; }
        public void Delete()
        {
            this.Column(ColumnNumber).Delete(XLShiftDeletedCells.ShiftCellsLeft);
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
                foreach (var c in Internals.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.Column == Internals.FirstCellAddress.Column))
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
            var address = new XLAddress(row, 1);
            return this.Range(address, address);
        }

        public IXLRange Column(int column)
        {
            return this;
        }

        public IXLRange Column(string column)
        {
            return this;
        }

        #endregion

        public IXLRangeInternals Internals { get; private set; }
    }
}
