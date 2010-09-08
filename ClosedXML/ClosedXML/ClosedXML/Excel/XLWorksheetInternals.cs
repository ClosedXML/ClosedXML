using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLWorksheetInternals: IXLWorksheetInternals
    {
        public XLWorksheetInternals(
            Dictionary<IXLAddress, IXLCell> cellsCollection , 
            Dictionary<Int32, IXLColumn> columnsCollection, 
            Dictionary<Int32, IXLRow> rowsCollection,
            List<String> mergedCells)
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedCells = mergedCells;
        }
        public IXLAddress FirstCellAddress
        {
            get { return new XLAddress(1, 1); }
        }

        public IXLAddress LastCellAddress
        {
            get { return new XLAddress(XLWorksheet.MaxNumberOfRows, XLWorksheet.MaxNumberOfColumns); }
        }
        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; private set; }
        public Dictionary<Int32, IXLColumn> ColumnsCollection { get; private set; }
        public Dictionary<Int32, IXLRow> RowsCollection { get; private set; }
        public List<String> MergedCells { get; private set; }
    }
}
