using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    
    internal delegate void RangeShiftedRowsDelegate(XLRange range, Int32 rowsShifted);
    internal delegate void RangeShiftedColumnsDelegate(XLRange range, Int32 columnsShifted);
    internal class XLWorksheet : XLRangeBase, IXLWorksheet
    {
        public event RangeShiftedRowsDelegate RangeShiftedRows;
        public event RangeShiftedColumnsDelegate RangeShiftedColumns;

        #region Constants

        public const Int32 MaxNumberOfRows = 1048576;
        public const Int32 MaxNumberOfColumns = 16384;

        #endregion

        public XLWorksheet(String sheetName, XLWorkbook workbook)
        {
            Worksheet = this;
            Style = workbook.Style;
            Internals = new XLWorksheetInternals(new Dictionary<IXLAddress, XLCell>(), new XLColumnsCollection(), new XLRowsCollection(), new List<String>());
            FirstAddressInSheet = new XLAddress(1, 1);
            LastAddressInSheet = new XLAddress(MaxNumberOfRows, MaxNumberOfColumns);
            PageSetup = new XLPageSetup(workbook.PageOptions, this);
            ColumnWidth = workbook.ColumnWidth;
            RowHeight = workbook.RowHeight;
            this.Name = sheetName;
            RangeShiftedRows += new RangeShiftedRowsDelegate(XLWorksheet_RangeShiftedRows);
            RangeShiftedColumns += new RangeShiftedColumnsDelegate(XLWorksheet_RangeShiftedColumns);
        }

        void XLWorksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            var newMerge = new List<String>();
            foreach (var merge in Internals.MergedCells)
            {
                var rng = Range(merge);
                if (range.FirstAddressInSheet.ColumnNumber <= rng.FirstAddressInSheet.ColumnNumber)
                {
                    var newRng = Range(
                        rng.FirstAddressInSheet.RowNumber,
                        rng.FirstAddressInSheet.ColumnNumber + columnsShifted,
                        rng.LastAddressInSheet.RowNumber,
                        rng.LastAddressInSheet.ColumnNumber + columnsShifted);
                    newMerge.Add(newRng.ToString());
                }
                else
                {
                    newMerge.Add(rng.ToString());
                }
            }
            Internals.MergedCells = newMerge;
        }

        void XLWorksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            var newMerge = new List<String>();
            foreach (var merge in Internals.MergedCells)
            {
                var rng = Range(merge);
                if (range.FirstAddressInSheet.RowNumber <= rng.FirstAddressInSheet.RowNumber)
                {
                    var newRng = Range(
                        rng.FirstAddressInSheet.RowNumber + rowsShifted,
                        rng.FirstAddressInSheet.ColumnNumber,
                        rng.LastAddressInSheet.RowNumber + rowsShifted,
                        rng.LastAddressInSheet.ColumnNumber);
                    newMerge.Add(newRng.ToString());
                }
                else
                {
                    newMerge.Add(rng.ToString());
                }
            }
            Internals.MergedCells = newMerge;
        }

        public void NotifyRangeShiftedRows(XLRange range, Int32 rowsShifted)
        {
            if (RangeShiftedRows != null)
                RangeShiftedRows(range, rowsShifted);
        }

        public void NotifyRangeShiftedColumns(XLRange range, Int32 columnsShifted)
        {
            if (RangeShiftedColumns != null)
                RangeShiftedColumns(range, columnsShifted);
        }

        public XLWorksheetInternals Internals { get; private set; }
        
        #region IXLStylized Members

        private IXLStyle style;
        public override IXLStyle Style
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

        public override IEnumerable<IXLStyle> Styles
        {
            get 
            {
                UpdatingStyle = true;
                yield return style;
                foreach (var c in Internals.CellsCollection.Values)
                {
                    yield return c.Style;
                }
                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        #endregion

        public Double ColumnWidth { get; set; }
        public Double RowHeight { get; set; }

        public String Name { get; set; }

        public IXLPageSetup PageSetup { get; private set; }

        public IXLRow FirstRowUsed()
        {
            var rngRow = this.AsRange().FirstRowUsed();
            if (rngRow != null)
            {
                return this.Row(rngRow.FirstAddressInSheet.RowNumber);
            }
            else
            {
                return null;
            }
        }
        public IXLRow LastRowUsed()
        {
            var rngRow = this.AsRange().LastRowUsed();
            if (rngRow != null)
            {
                return this.Row(rngRow.LastAddressInSheet.RowNumber);
            }
            else
            {
                return null;
            }
        }
        public IXLColumn FirstColumnUsed()
        {
            var rngColumn = this.AsRange().FirstColumnUsed();
            if (rngColumn != null)
            {
                return this.Column(rngColumn.FirstAddressInSheet.ColumnNumber);
            }
            else
            {
                return null;
            }
        }
        public IXLColumn LastColumnUsed()
        {
            var rngColumn = this.AsRange().LastColumnUsed();
            if (rngColumn != null)
            {
                return this.Column(rngColumn.LastAddressInSheet.ColumnNumber);
            }
            else
            {
                return null;
            }
        }

        public IXLColumns Columns()
        {
            var retVal = new XLColumns(this, true);
            var columnList = new List<Int32>();

            if (this.Internals.CellsCollection.Count > 0)
                columnList.AddRange(this.Internals.CellsCollection.Keys.Select(k => k.ColumnNumber).Distinct());

            if (this.Internals.ColumnsCollection.Count > 0)
                columnList.AddRange(this.Internals.ColumnsCollection.Keys.Where(c => !columnList.Contains(c)));

            foreach (var c in columnList)
            {
                retVal.Add((XLColumn)this.Column(c));
            }

            return retVal;
        }
        public IXLColumns Columns( String columns)
        {
            var retVal = new XLColumns(this);
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
                        retVal.Add((XLColumn)col);
                    }
                else
                    foreach (var col in this.Columns(firstColumn, lastColumn))
                    {
                        retVal.Add((XLColumn)col);
                    }
            }
            return retVal;
        }
        public IXLColumns Columns( String firstColumn, String lastColumn)
        {
            return this.Columns(XLAddress.GetColumnNumberFromLetter(firstColumn), XLAddress.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLColumns Columns( Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLColumns(this);

            for (var co = firstColumn; co <= lastColumn; co++)
            {
                retVal.Add((XLColumn)this.Column(co));
            }
            return retVal;
        }

        public IXLRows Rows()
        {
            var retVal = new XLRows(this, true);
            var rowList = new List<Int32>();

            if (this.Internals.CellsCollection.Count > 0)
                rowList.AddRange(this.Internals.CellsCollection.Keys.Select(k => k.RowNumber).Distinct());

            if (this.Internals.ColumnsCollection.Count > 0)
                rowList.AddRange(this.Internals.ColumnsCollection.Keys.Where(r => !rowList.Contains(r)));

            foreach (var r in rowList)
            {
                retVal.Add((XLRow)this.Row(r));
            }

            return retVal;
        }
        public IXLRows Rows( String rows)
        {
            var retVal = new XLRows(this);
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
                    retVal.Add((XLRow)row);
                }
            }
            return retVal;
        }
        public IXLRows Rows( Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRows(this);

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add((XLRow)this.Row(ro));
            }
            return retVal;
        }

        public IXLRow Row( Int32 row)
        {
            IXLStyle styleToUse;
            if (this.Internals.RowsCollection.ContainsKey(row))
            {
                styleToUse = this.Internals.RowsCollection[row].Style;
            }
            else
            {
                styleToUse = this.Style;
                this.Internals.RowsCollection.Add(row, new XLRow(row, new XLRowParameters(this, styleToUse, false)));
            }

            return new XLRow(row, new XLRowParameters(this, styleToUse, true));
        }
        public IXLColumn Column( Int32 column)
        {
            IXLStyle styleToUse;
            if (this.Internals.ColumnsCollection.ContainsKey(column))
            {
                styleToUse = this.Internals.ColumnsCollection[column].Style;
            }
            else
            {
                styleToUse = this.Style;
                this.Internals.ColumnsCollection.Add(column, new XLColumn(column, new XLColumnParameters(this, this.Style, false)));
            }

            return new XLColumn(column, new XLColumnParameters(this, this.Style, true));
        }
        public IXLColumn Column( String column)
        {
            return this.Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, XLWorksheet.MaxNumberOfRows, XLWorksheet.MaxNumberOfColumns);
        }

    }
}
