using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public class XLWorksheet: IXLWorksheet
    {
        #region Constants

        public const Int32 MaxNumberOfRows = 1048576;
        public const Int32 MaxNumberOfColumns = 16384;

        #endregion

        Dictionary<IXLAddress, IXLCell> cellsCollection = new Dictionary<IXLAddress, IXLCell>();

        public XLWorksheet(String sheetName)
        { 
            var defaultAddress = new XLAddress(0,0);
            DefaultCell = new XLCell(defaultAddress, XLWorkbook.DefaultStyle);
            cellsCollection.Add(defaultAddress, DefaultCell);
            this.Name = sheetName;
        }

        #region IXLRange Members

        public Dictionary<IXLAddress, IXLCell> CellsCollection
        {
            get { return cellsCollection; }
        }

        public IXLAddress FirstCellAddress
        {
            get { return new XLAddress(1, 1); }
        }

        public IXLAddress LastCellAddress
        {
            get { return new XLAddress(MaxNumberOfRows, MaxNumberOfColumns); }
        }

        private IXLCell DefaultCell { get; set; }

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
                foreach (var c in cellsCollection.Values)
                {
                    yield return c.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion

        public IXLRange Row(Int32 row)
        {
            return new XLRow(row, cellsCollection, Style);
        }
        public IXLColumn Column(Int32 column)
        {
            return new XLColumn(column, cellsCollection, Style);
        }
        public IXLColumn Column(String column)
        {
            return new XLColumn(XLAddress.GetColumnNumberFromLetter(column), cellsCollection, Style);
        }

        #region IXLRange Members

        
        IXLRange IXLRange.Column(int column)
        {
            IXLAddress firstCellAddress = new XLAddress(1, column);
            IXLAddress lastCellAddress = new XLAddress(MaxNumberOfRows, column);
            return this.Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLRange.Column(string column)
        {
            IXLAddress firstCellAddress = new XLAddress(1, column);
            IXLAddress lastCellAddress = new XLAddress(MaxNumberOfRows, column);
            return this.Range(firstCellAddress, lastCellAddress);
        }

        #endregion


        public String Name { get; set; }
    }
}
