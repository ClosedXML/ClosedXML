using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLRow: IXLRange
    {
        public XLRow(Int32 row, Dictionary<IXLAddress, IXLCell> cellsCollection, IXLStyle defaultStyle)
        {
            FirstCellAddress = new XLAddress(row, 1);
            LastCellAddress = new XLAddress(row, XLWorksheet.MaxNumberOfColumns); 
            CellsCollection = cellsCollection;

            var defaultAddress = new XLAddress(row, 0);
            if (!cellsCollection.ContainsKey(defaultAddress))
            {
                DefaultCell = new XLCell(defaultAddress, defaultStyle);
                cellsCollection.Add(defaultAddress, DefaultCell);
            }
            else
            {
                DefaultCell = cellsCollection[defaultAddress];
            }
        }

        private IXLCell DefaultCell { get; set; }

        #region IXLRange Members

        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; private set; }
        public IXLAddress FirstCellAddress { get; private set; }
        public IXLAddress LastCellAddress { get; private set; }

        #endregion

        #region IXLStylized Members

        public IXLStyle Style
        {
            get
            {
                return DefaultCell.Style;
            }
            set
            {
                DefaultCell.Style = value;
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
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
            var address = new XLAddress(FirstCellAddress.Row, column);
            return this.Range(address, address);
        }

        public IXLRange Column(string column)
        {
            var address = new XLAddress(FirstCellAddress.Row, column);
            return this.Range(address, address);
        }

        #endregion
    }
}
