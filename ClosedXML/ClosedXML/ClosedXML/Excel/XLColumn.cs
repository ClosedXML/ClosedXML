using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLColumn: IXLColumn
    {
        public XLColumn(Int32 column, Dictionary<IXLAddress, IXLCell> cellsCollection, IXLStyle defaultStyle)
        {
            FirstCellAddress = new XLAddress(1, column);
            LastCellAddress = new XLAddress(XLWorksheet.MaxNumberOfRows, column);
            RowNumber = 1;
            ColumnNumber = column;
            ColumnLetter = XLAddress.GetColumnLetterFromNumber(column);
            CellsCollection = cellsCollection;
            this.style = new XLStyle(this, defaultStyle);
            this.Width = XLWorkbook.DefaultColumnWidth;
        }

        public Double Width { get; set; }
        public Int32 RowNumber { get; private set; }
        public Int32 ColumnNumber { get; private set; }
        public String ColumnLetter { get; private set; }

        #region IXLRange Members

        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; private set; }
        public List<String> MergedCells { get; private set; }
        public IXLAddress FirstCellAddress { get; private set; }
        public IXLAddress LastCellAddress { get; private set; }

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
                foreach (var c in CellsCollection.Values.Where(c => c.Address.Column == FirstCellAddress.Column))
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
    }
}
