using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals: IXLWorksheetInternals
    {
        public XLWorksheetInternals(
            Dictionary<IXLAddress, XLCell> cellsCollection , 
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            List<String> mergedCells)
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedCells = mergedCells;
        }

        public Dictionary<IXLAddress, XLCell> CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public List<String> MergedCells { get; internal set; }
    }
}
