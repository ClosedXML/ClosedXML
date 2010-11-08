using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRangeRow: XLRangeBase, IXLRangeRow
    {
        public XLRangeRow(XLRangeParameters xlRangeParameters): base(xlRangeParameters.RangeAddress)
        {
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.RangeAddress.FirstAddress.RowNumber <= RangeAddress.FirstAddress.RowNumber
                && range.RangeAddress.LastAddress.RowNumber >= RangeAddress.LastAddress.RowNumber)
            {
                if (range.RangeAddress.FirstAddress.ColumnNumber <= RangeAddress.FirstAddress.ColumnNumber)
                    RangeAddress.FirstAddress = new XLAddress(RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + columnsShifted);

                if (range.RangeAddress.FirstAddress.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber)
                    RangeAddress.LastAddress = new XLAddress(RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber + columnsShifted);
            }
        }
        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.FirstAddress.ColumnNumber <= RangeAddress.FirstAddress.ColumnNumber
                && range.RangeAddress.LastAddress.ColumnNumber >= RangeAddress.LastAddress.ColumnNumber)
            {
                if (range.RangeAddress.FirstAddress.RowNumber <= RangeAddress.FirstAddress.RowNumber)
                    RangeAddress.FirstAddress = new XLAddress(RangeAddress.FirstAddress.RowNumber + rowsShifted, RangeAddress.FirstAddress.ColumnNumber);

                if (range.RangeAddress.FirstAddress.RowNumber <= RangeAddress.LastAddress.RowNumber)
                    RangeAddress.LastAddress = new XLAddress(RangeAddress.LastAddress.RowNumber + rowsShifted, RangeAddress.LastAddress.ColumnNumber);
            }
        }

        public IXLCell Cell(int column)
        {
            return Cell(1, column);
        }
        public new IXLCell Cell(string column)
        {
            return Cell(1, column);
        }

        public IEnumerable<IXLCell> Cells(int firstColumn, int lastColumn)
        {
            return Cells()
                .Where(c => c.Address.ColumnNumber >= firstColumn
                    && c.Address.ColumnNumber <= lastColumn);
        }
        public IEnumerable<IXLCell> Cells(String firstColumn, String lastColumn)
        {
            return Cells()
                .Where(c => c.Address.ColumnNumber >= XLAddress.GetColumnNumberFromLetter(firstColumn)
                    && c.Address.ColumnNumber <= XLAddress.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRange Range(int firstColumn, int lastColumn)
        {
            return Range(1, firstColumn, 1, lastColumn);
        }
    }
}

