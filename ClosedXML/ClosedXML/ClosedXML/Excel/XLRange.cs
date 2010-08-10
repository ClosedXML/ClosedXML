using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLRange: IXLRange
    {
        private IXLStyle defaultStyle;
        public XLRange(IXLAddress firstCellAddress, IXLAddress lastCellAddress, Dictionary<IXLAddress, IXLCell> cellCollection)
        {
            FirstCellAddress = firstCellAddress;
            LastCellAddress = lastCellAddress;
            CellsCollection = cellCollection;
            this.defaultStyle = new XLStyle(this, this.FirstCell().Style);
        }

        #region IXLRange Members

        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; private set; }

        public IXLAddress FirstCellAddress { get; private set; }

        public IXLAddress LastCellAddress { get; private set; }

        public IXLRange Row(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(row, 1);
            IXLAddress lastCellAddress = new XLAddress(row, this.ColumnCount());
            return this.Range(firstCellAddress, lastCellAddress);
        }
        public IXLRange Column(Int32 column)
        {
            IXLAddress firstCellAddress = new XLAddress(1, column);
            IXLAddress lastCellAddress = new XLAddress(this.RowCount(), column);
            return this.Range(firstCellAddress, lastCellAddress);
        }
        public IXLRange Column(String column)
        {
            IXLAddress firstCellAddress = new XLAddress(1, column);
            IXLAddress lastCellAddress = new XLAddress(this.RowCount(), column);
            return this.Range(firstCellAddress, lastCellAddress);
        }


        #endregion

        #region IXLStylized Members

        public IXLStyle Style 
        {
            get
            {
                return this.defaultStyle;
            }
            set
            {
                this.Cells().ForEach(c => c.Style = value);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                foreach (var cell in this.Cells())
                {
                    yield return cell.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion
    }
}
