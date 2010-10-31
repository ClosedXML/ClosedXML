using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRange: XLRangeBase, IXLRange
    {
        public XLRange(XLRangeParameters xlRangeParameters)
        {
            FirstAddressInSheet = xlRangeParameters.FirstCellAddress;
            LastAddressInSheet = xlRangeParameters.LastCellAddress;
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            //Worksheet.Internals.RowsCollection.RowShifted += new RowShiftedDelegate(RowsCollection_RowShifted);
            //Worksheet.Internals.ColumnsCollection.ColumnShifted += new ColumnShiftedDelegate(ColumnsCollection_ColumnShifted);
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.FirstAddressInSheet.RowNumber <= FirstAddressInSheet.RowNumber
                && range.LastAddressInSheet.RowNumber >= LastAddressInSheet.RowNumber)
            {
                if (range.FirstAddressInSheet.ColumnNumber <= FirstAddressInSheet.ColumnNumber)
                    FirstAddressInSheet = new XLAddress(FirstAddressInSheet.RowNumber, FirstAddressInSheet.ColumnNumber + columnsShifted);

                if (range.FirstAddressInSheet.ColumnNumber <= LastAddressInSheet.ColumnNumber)
                    LastAddressInSheet = new XLAddress(LastAddressInSheet.RowNumber, LastAddressInSheet.ColumnNumber + columnsShifted);
            }
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.FirstAddressInSheet.ColumnNumber <= FirstAddressInSheet.ColumnNumber
                && range.LastAddressInSheet.ColumnNumber >= LastAddressInSheet.ColumnNumber)
            {
                if (range.FirstAddressInSheet.RowNumber <= FirstAddressInSheet.RowNumber)
                    FirstAddressInSheet = new XLAddress(FirstAddressInSheet.RowNumber + rowsShifted, FirstAddressInSheet.ColumnNumber);

                if (range.FirstAddressInSheet.RowNumber <= LastAddressInSheet.RowNumber)
                    LastAddressInSheet = new XLAddress(LastAddressInSheet.RowNumber + rowsShifted, LastAddressInSheet.ColumnNumber);
            }
        }

        #region IXLRange Members

        public IXLRangeColumn FirstColumn()
        {
            return this.Column(1);
        }
        public IXLRangeColumn LastColumn()
        {
            return this.Column(this.ColumnCount());
        }
        public IXLRangeColumn FirstColumnUsed()
        {
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 minColumnUsed = Int32.MaxValue;
            Int32 minColumnInCells = Int32.MaxValue;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                minColumnInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Min();

            Int32 minCoInColumns = Int32.MaxValue;
            if (this.Worksheet.Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                minCoInColumns = this.Worksheet.Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Min();

            minColumnUsed = minColumnInCells < minCoInColumns ? minColumnInCells : minCoInColumns;

            if (minColumnUsed == Int32.MaxValue)
                return null;
            else
                return this.Column(minColumnUsed);
        }
        public IXLRangeColumn LastColumnUsed()
        {
            var firstColumn = this.FirstAddressInSheet.ColumnNumber;
            var columnCount = this.ColumnCount();
            Int32 maxColumnUsed = 0;
            Int32 maxColumnInCells = 0;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount))
                maxColumnInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.ColumnNumber >= firstColumn && c.Key.ColumnNumber <= columnCount).Select(c => c.Key.ColumnNumber).Max();

            Int32 maxCoInColumns = 0;
            if (this.Worksheet.Internals.ColumnsCollection.Any(c => c.Key >= firstColumn && c.Key <= columnCount))
                maxCoInColumns = this.Worksheet.Internals.ColumnsCollection
                    .Where(c => c.Key >= firstColumn && c.Key <= columnCount).Select(c => c.Key).Max();

            maxColumnUsed = maxColumnInCells > maxCoInColumns ? maxColumnInCells : maxCoInColumns;

            if (maxColumnUsed == 0)
                return null;
            else
                return this.Column(maxColumnUsed);
        }

        public IXLRangeRow FirstRow()
        {
            return this.Row(1);
        }
        public IXLRangeRow LastRow()
        {
            return this.Row(this.RowCount());
        }
        public IXLRangeRow FirstRowUsed()
        {
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var rowCount = this.RowCount();
            Int32 minRowUsed = Int32.MaxValue;
            Int32 minRowInCells = Int32.MaxValue;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                minRowInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Min();

            Int32 minRoInRows = Int32.MaxValue;
            if (this.Worksheet.Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                minRoInRows = this.Worksheet.Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Min();

            minRowUsed = minRowInCells < minRoInRows ? minRowInCells : minRoInRows;

            if (minRowUsed == Int32.MaxValue)
                return null;
            else
                return this.Row(minRowUsed);
        }
        public IXLRangeRow LastRowUsed()
        {
            var firstRow = this.FirstAddressInSheet.RowNumber;
            var rowCount = this.RowCount();
            Int32 maxRowUsed = 0;
            Int32 maxRowInCells = 0;
            if (this.Worksheet.Internals.CellsCollection.Any(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount))
                maxRowInCells = this.Worksheet.Internals.CellsCollection
                    .Where(c => c.Key.RowNumber >= firstRow && c.Key.RowNumber <= rowCount).Select(c => c.Key.RowNumber).Max();

            Int32 maxRoInRows = 0;
            if (this.Worksheet.Internals.RowsCollection.Any(r => r.Key >= firstRow && r.Key <= rowCount))
                maxRoInRows = this.Worksheet.Internals.RowsCollection
                    .Where(r => r.Key >= firstRow && r.Key <= rowCount).Select(r => r.Key).Max();

            maxRowUsed = maxRowInCells > maxRoInRows ? maxRowInCells : maxRoInRows;

            if (maxRowUsed == 0)
                return null;
            else
                return this.Row(maxRowUsed);
        }

        public IXLRangeRow Row(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(FirstAddressInSheet.RowNumber + row - 1, FirstAddressInSheet.ColumnNumber);
            IXLAddress lastCellAddress = new XLAddress(FirstAddressInSheet.RowNumber + row - 1, LastAddressInSheet.ColumnNumber);
            return new XLRangeRow(
                new XLRangeParameters(
                    firstCellAddress, 
                    lastCellAddress, 
                    Worksheet, 
                    Worksheet.Style));
                
        }
        public IXLRangeColumn Column(Int32 column)
        {
            IXLAddress firstCellAddress = new XLAddress(FirstAddressInSheet.RowNumber, FirstAddressInSheet.ColumnNumber + column - 1);
            IXLAddress lastCellAddress = new XLAddress(LastAddressInSheet.RowNumber, FirstAddressInSheet.ColumnNumber + column - 1);
            return new XLRangeColumn(
                new XLRangeParameters(
                    firstCellAddress,
                    lastCellAddress,
                    Worksheet,
                    Worksheet.Style));
        }
        public IXLRangeColumn Column(String column)
        {
            return this.Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        public IXLRangeColumns Columns()
        {
            var retVal = new XLRangeColumns(Worksheet);
            foreach (var c in Enumerable.Range(1, this.ColumnCount()))
            {
                retVal.Add(this.Column(c));
            }
            return retVal;
        }
        public IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLRangeColumns(Worksheet);

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add(this.Column(co));
            }
            return retVal;
        }
        public IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return this.Columns(XLAddress.GetColumnNumberFromLetter(firstColumn), XLAddress.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns(Worksheet);
            var columnPairs = columns.Split(',');
            foreach (var pair in columnPairs)
            {
                String firstColumn;
                String lastColumn;
                if (pair.Contains(':'))
                {
                    var columnRange = pair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = pair;
                    lastColumn = pair;
                }

                Int32 tmp;
                if (Int32.TryParse(firstColumn, out tmp))
                    foreach (var col in this.Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                    {
                        retVal.Add(col);
                    }
                else
                    foreach (var col in this.Columns(firstColumn, lastColumn))
                    {
                        retVal.Add(col);
                    }
            }
            return retVal;
        }

        public IXLRangeRows Rows()
        {
            var retVal = new XLRangeRows(Worksheet);
            foreach (var r in Enumerable.Range(1, this.RowCount()))
            {
                retVal.Add(this.Row(r));
            }
            return retVal;
        }
        public IXLRangeRows Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRangeRows(Worksheet);

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(this.Row(ro));
            }
            return retVal;
        }
        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows(Worksheet);
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
            {
                String firstRow;
                String lastRow;
                if (pair.Contains(':'))
                {
                    var rowRange = pair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = pair;
                    lastRow = pair;
                }
                foreach (var row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        #endregion

    }
}
