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
            Style = XLWorkbook.DefaultStyle;
            MergedCells = new List<String>();
            RowNumber = 1;
            ColumnNumber = 1;
            ColumnLetter = "A";
            PrintOptions = new XLPrintOptions();
            this.Name = sheetName;
        }

        public IXLPrintOptions PrintOptions { get; private set; }

        #region IXLRange Members
        public List<String> MergedCells { get; private set; }
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

        public Int32 RowNumber { get; private set; }
        public Int32 ColumnNumber { get; private set; }
        public String ColumnLetter { get; private set; }

        public List<IXLColumn> Columns()
        {
            var retVal = new List<IXLColumn>();
            var columnList = new List<Int32>();

            if (CellsCollection.Count > 0)
            columnList.AddRange(CellsCollection.Keys.Select(k => k.Column).Distinct());

            if (ColumnsCollection.Count > 0)
            columnList.AddRange(ColumnsCollection.Keys.Where(c => !columnList.Contains(c)));

            foreach (var c in columnList)
            {
                retVal.Add(Column(c));
            }

            return retVal;
        }
        
        public IXLRange PrintArea { get; set; }

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
                var xlRangeParameters = new XLRangeParameters()
                {
                    CellsCollection = cellsCollection,
                    MergedCells = this.MergedCells,
                    DefaultStyle = Style,
                    PrintArea = this.PrintArea
                };
                xlRow = new XLRow(row, xlRangeParameters);
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
                var xlRangeParameters = new XLRangeParameters()
                {
                    CellsCollection = cellsCollection,
                    MergedCells = this.MergedCells,
                    DefaultStyle = Style,
                    PrintArea = this.PrintArea
                };
                xlColumn = new XLColumn(column, xlRangeParameters);
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
