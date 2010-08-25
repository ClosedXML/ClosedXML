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
        Dictionary<Int32, IXLRow> rowsCollection = new Dictionary<Int32, IXLRow>();
        Dictionary<Int32, IXLColumn> columnsCollection = new Dictionary<Int32, IXLColumn>();

        public XLWorksheet(String sheetName)
        { 
            var defaultAddress = new XLAddress(0,0);
            DefaultCell = new XLCell(defaultAddress, XLWorkbook.DefaultStyle);
            cellsCollection.Add(defaultAddress, DefaultCell);
            var tmp = this.Cell(1, 1).Value;
            this.Name = sheetName;
        }

        #region IXLRange Members

        public Dictionary<IXLAddress, IXLCell> CellsCollection
        {
            get { return cellsCollection; }
        }
        public Dictionary<Int32, IXLColumn> ColumnsCollection
        {
            get { return columnsCollection; }
        }
        public Dictionary<Int32, IXLRow> RowsCollection
        {
            get { return rowsCollection; }
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

        public IXLRow Row(Int32 row)
        {
            IXLRow xlRow;
            if (rowsCollection.ContainsKey(row))
            {
                xlRow = rowsCollection[row];
            }
            else
            {
                xlRow = new XLRow(row, cellsCollection, Style);
                rowsCollection.Add(row, xlRow);
            }

            return xlRow;
        }
        public IXLColumn Column(Int32 column)
        {
            IXLColumn xlColumn;
            if (columnsCollection.ContainsKey(column))
            {
                xlColumn = columnsCollection[column];
            }
            else
            {
                xlColumn = new XLColumn(column, cellsCollection, Style);
                columnsCollection.Add(column, xlColumn);
            }

            return xlColumn;
        }
        public IXLColumn Column(String column)
        {
            return Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        #region IXLRange Members

        IXLRange IXLRange.Row(Int32 row)
        {
            var firstCellAddress = new XLAddress(row, 1);
            var lastCellAddress = new XLAddress(row, MaxNumberOfColumns);
            return this.Range(firstCellAddress, lastCellAddress);
        }
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
