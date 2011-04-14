using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRange: XLRangeBase, IXLRange
    {
        //public new IXLWorksheet Worksheet { get { return base.Worksheet; } }
        public XLRangeParameters RangeParameters { get; private set; }
        public XLRange(XLRangeParameters xlRangeParameters): base(xlRangeParameters.RangeAddress)
        {
            this.RangeParameters = xlRangeParameters;
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            ShiftColumns(this.RangeAddress, range, columnsShifted);
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            ShiftRows(this.RangeAddress, range, rowsShifted);
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
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
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
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
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
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
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
            var firstRow = this.RangeAddress.FirstAddress.RowNumber;
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
            IXLAddress firstCellAddress = new XLAddress(RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.FirstAddress.ColumnNumber, false, false);
            IXLAddress lastCellAddress = new XLAddress(RangeAddress.FirstAddress.RowNumber + row - 1, RangeAddress.LastAddress.ColumnNumber, false, false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), 
                    Worksheet, 
                    Worksheet.Style));
                
        }
        public IXLRangeColumn Column(Int32 column)
        {
            IXLAddress firstCellAddress = new XLAddress(RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            IXLAddress lastCellAddress = new XLAddress(RangeAddress.LastAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + column - 1, false, false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress),
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
                var tPair = pair.Trim();
                String firstColumn;
                String lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                        tPair = tPair.Replace('-', ':');

                    var columnRange = tPair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
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
                var tPair = pair.Trim();
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                        tPair = tPair.Replace('-', ':');

                    var rowRange = tPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        public void Transpose(XLTransposeOptions transposeOption)
        {
            var rowCount = this.RowCount();
            var columnCount = this.ColumnCount();
            var squareSide = rowCount > columnCount ? rowCount : columnCount;

            var firstCell = FirstCell();
            var lastCell = LastCell();

            MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
            TransposeMerged(squareSide);
            TransposeRange(squareSide);
            this.RangeAddress.LastAddress = new XLAddress(
                firstCell.Address.RowNumber + columnCount - 1,
                firstCell.Address.ColumnNumber + rowCount - 1,
                RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
            if (rowCount > columnCount)
            {
                var rng = Worksheet.Range(
                    this.RangeAddress.LastAddress.RowNumber + 1,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                    this.RangeAddress.LastAddress.ColumnNumber);
                rng.Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
            else if (columnCount > rowCount)
            {
                var rng = Worksheet.Range(
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + 1,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
                rng.Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }

            foreach (var c in this.Range(1,1,columnCount, rowCount).Cells())
            {
                var border = new XLBorder(this, c.Style.Border);
                c.Style.Border.TopBorder = border.LeftBorder;
                c.Style.Border.TopBorderColor = border.LeftBorderColor;
                c.Style.Border.LeftBorder = border.TopBorder;
                c.Style.Border.LeftBorderColor = border.TopBorderColor;
                c.Style.Border.RightBorder = border.BottomBorder;
                c.Style.Border.RightBorderColor = border.BottomBorderColor;
                c.Style.Border.BottomBorder = border.RightBorder;
                c.Style.Border.BottomBorderColor = border.RightBorderColor;
            }
        }

        private void TransposeRange(int squareSide)
        {
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            XLRange rngToTranspose = (XLRange)Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                this.RangeAddress.FirstAddress.ColumnNumber,
                this.RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                this.RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            Int32 roInitial = rngToTranspose.RangeAddress.FirstAddress.RowNumber;
            Int32 coInitial = rngToTranspose.RangeAddress.FirstAddress.ColumnNumber;
            Int32 roCount = rngToTranspose.RowCount();
            Int32 coCount = rngToTranspose.ColumnCount();
            for (Int32 ro = 1; ro <= roCount; ro++)
            {
                for (Int32 co = 1; co <= coCount; co++)
                {
                    var oldCell = rngToTranspose.Cell(ro, co);
                    var newKey = rngToTranspose.Cell(co, ro).Address; // new XLAddress(c.Address.ColumnNumber, c.Address.RowNumber);
                    var newCell = new XLCell(newKey, oldCell.Style, Worksheet);
                    newCell.Value = oldCell.Value;
                    newCell.DataType = oldCell.DataType;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(oldCell.Address);
                }
            }
            //foreach (var c in rngToTranspose.Cells())
            //{
            //    var newKey = new XLAddress(c.Address.ColumnNumber, c.Address.RowNumber);
            //    var newCell = new XLCell(newKey, c.Style, Worksheet);
            //    newCell.Value = c.Value;
            //    newCell.DataType = c.DataType;
            //    cellsToInsert.Add(newKey, newCell);
            //    cellsToDelete.Add(c.Address);
            //}
            cellsToDelete.ForEach(c => this.Worksheet.Internals.CellsCollection.Remove(c));
            cellsToInsert.ForEach(c => this.Worksheet.Internals.CellsCollection.Add(c.Key, c.Value));
        }

        private void TransposeMerged(Int32 squareSide)
        {
            XLRange rngToTranspose = (XLRange)Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                this.RangeAddress.FirstAddress.ColumnNumber,
                this.RangeAddress.FirstAddress.RowNumber + squareSide - 1,
                this.RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

            List<IXLRange> mergeToDelete = new List<IXLRange>();
            List<IXLRange> mergeToInsert = new List<IXLRange>();
            foreach (var merge in Worksheet.Internals.MergedRanges)
            {
                if (this.Contains(merge))
                {
                    merge.RangeAddress.LastAddress = rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address;
                }
            }
            mergeToDelete.ForEach(m => this.Worksheet.Internals.MergedRanges.Remove(m));
            mergeToInsert.ForEach(m => this.Worksheet.Internals.MergedRanges.Add(m));
        }

        private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
        {
            if (transposeOption == XLTransposeOptions.MoveCells)
            {
                if (rowCount > columnCount)
                {
                    this.InsertColumnsAfter(rowCount - columnCount, false);
                }
                else if (columnCount > rowCount)
                {
                    this.InsertRowsBelow(columnCount - rowCount, false);
                }
            }
            else
            {
                if (rowCount > columnCount)
                {
                    var toMove = rowCount - columnCount;
                    var rngToClear = Worksheet.Range(
                        this.RangeAddress.FirstAddress.RowNumber,
                        this.RangeAddress.LastAddress.ColumnNumber + 1,
                        this.RangeAddress.LastAddress.RowNumber,
                        this.RangeAddress.LastAddress.ColumnNumber + toMove);
                    rngToClear.Clear();
                }
                else if (columnCount > rowCount)
                {
                    var toMove = columnCount - rowCount;
                    var rngToClear = Worksheet.Range(
                        this.RangeAddress.LastAddress.RowNumber + 1,
                        this.RangeAddress.FirstAddress.ColumnNumber,
                        this.RangeAddress.LastAddress.RowNumber + toMove,
                        this.RangeAddress.LastAddress.ColumnNumber);
                    rngToClear.Clear();
                }
            }
        }

        public IXLTable AsTable()
        {
            return new XLTable(this, false);
        }

        public IXLTable AsTable(String name)
        {
            return new XLTable(this, name, false);
        }

        public IXLTable CreateTable()
        {
            return new XLTable(this, true);
        }

        public IXLTable CreateTable(String name)
        {
            return new XLTable(this, name, true);
        }
        #endregion

        public override bool Equals(object obj)
        {
            var other = (XLRange)obj;
            return this.RangeAddress.Equals(other.RangeAddress)
                && this.Worksheet.Equals(other.Worksheet);
        }

        public override int GetHashCode()
        {
            return RangeAddress.GetHashCode()
                    ^ this.Worksheet.GetHashCode();
        }

        private void SortRange(XLRange xLRange, string[] columns)
        {
            throw new NotImplementedException();
        }

        public IXLRange Sort(String columnsToSort)
        {
            var cols = columnsToSort.Split(',').ToList();
            q_sort(1, this.RowCount(), cols);
            return this;
        }

        public void q_sort(int left, int right, List<String> columnsToSort)
        {
            int i, j;
            XLRangeRow x, y;

            i = left; j = right;
            x = (XLRangeRow)Row(((left + right) / 2));

            do
            {
                while ((((XLRangeRow)Row(i)).CompareTo(x, columnsToSort) < 0) && (i < right)) i++;
                while ((0 < ((XLRangeRow)Row(j)).CompareTo(x, columnsToSort)) && (j > left)) j--;

                if (i < j)
                {
                    SwapRows(i, j);
                    i++; j--;
                }
                else if (i == j)
                {
                    i++; j--;
                }
            } while (i <= j);

            if (left < j) q_sort(left, j, columnsToSort);
            if (i < right) q_sort(i, right, columnsToSort);

        }

        public void SwapRows(Int32 row1, Int32 row2)
        {

            Int32 cellCount = ColumnCount();

            for (Int32 co = 1; co <= cellCount; co++)
            {

                var cell1 = (XLCell)Row(row1).Cell(co);
                var cell1Address = cell1.Address;
                var cell2 = (XLCell)Row(row2).Cell(co);

                cell1.Address = cell2.Address;

                cell2.Address = cell1Address;
                Worksheet.Internals.CellsCollection.Remove(cell1.Address);
                Worksheet.Internals.CellsCollection.Remove(cell2.Address);
                Worksheet.Internals.CellsCollection.Add(cell1.Address, cell1);
                Worksheet.Internals.CellsCollection.Add(cell2.Address, cell2);

            }

        }
    }
}
