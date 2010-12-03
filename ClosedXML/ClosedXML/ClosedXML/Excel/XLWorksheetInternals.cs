using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheetInternals: IXLWorksheetInternals
    {
        public XLWorksheetInternals(
            XLCellCollection cellsCollection, 
            XLColumnsCollection columnsCollection,
            XLRowsCollection rowsCollection,
            List<String> mergedCells,
            XLWorkbook workbook
            )
        {
            CellsCollection = cellsCollection;
            ColumnsCollection = columnsCollection;
            RowsCollection = rowsCollection;
            MergedCells = mergedCells;
            Workbook = workbook;
        }

        public XLCellCollection CellsCollection { get; private set; }
        public XLColumnsCollection ColumnsCollection { get; private set; }
        public XLRowsCollection RowsCollection { get; private set; }
        public List<String> MergedCells { get; internal set; }
        public XLWorkbook Workbook { get; internal set; }
    }
}
