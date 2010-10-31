using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRangeRow: XLRangeBase, IXLRangeRow
    {
        public XLRangeRow(XLRangeParameters xlRangeParameters)
        {
            FirstAddressInSheet = xlRangeParameters.FirstCellAddress;
            LastAddressInSheet = xlRangeParameters.LastCellAddress;
            Worksheet = xlRangeParameters.Worksheet;
            Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
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

