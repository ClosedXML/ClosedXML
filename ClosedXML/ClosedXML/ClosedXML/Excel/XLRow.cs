using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLRow: IXLRow
    {
        public XLRow(Int32 row, Dictionary<IXLAddress, IXLCell> cellsCollection, IXLStyle defaultStyle)
        {
            FirstCellAddress = new XLAddress(row, 1);
            LastCellAddress = new XLAddress(row, XLWorksheet.MaxNumberOfColumns);
            RowNumber = row;
            ColumnNumber = 1;
            ColumnLetter = "A";
            CellsCollection = cellsCollection;
            this.style = new XLStyle(this, defaultStyle);
            this.Height = XLWorkbook.DefaultRowHeight;
        }

        public Double Height { get; set; }
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
                foreach (var c in CellsCollection.Values.Where(c => c.Address.Row == FirstCellAddress.Row))
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
