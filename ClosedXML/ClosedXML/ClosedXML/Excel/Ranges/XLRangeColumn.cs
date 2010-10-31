using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRangeColumn: XLRangeBase, IXLRangeColumn
    {
        public XLRangeColumn(XLRangeParameters xlRangeParameters)
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

        public IXLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        public IEnumerable<IXLCell> Cells(int firstRow, int lastRow)
        {
            return Cells()
                .Where(c => c.Address.RowNumber >= firstRow
                    && c.Address.RowNumber <= lastRow);
        }


        public IXLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
        }
    }
}

