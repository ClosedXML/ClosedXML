using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals: IXLWorksheetInternals
    {
        public XLWorksheetInternals(
            Dictionary<IXLAddress, IXLCell> cellsCollection , 
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            List<String> mergedCells)
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedCells = mergedCells;
        }

        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public List<String> MergedCells { get; internal set; }
    }
}
