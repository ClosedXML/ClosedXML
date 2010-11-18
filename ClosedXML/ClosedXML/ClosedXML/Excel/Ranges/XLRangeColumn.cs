using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRangeColumn: XLRangeBase, IXLRangeColumn
    {
        public XLRangeColumn(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters.RangeAddress)
        {
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

