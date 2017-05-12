using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Drawings
{
    public class XLMarker : IXLMarker
    {
        private Int32 colId;
        private Int32 rowId;
        private Double colOffset;
        private Double rowOffset;

        public Int32 ColumnId
        {
            set
            {
                if (value < 1 || value > XLHelper.MaxColumnNumber)
                    throw new ArgumentOutOfRangeException(String.Format("Column number must be between 1 and {0}",
                                                                 XLHelper.MaxColumnNumber));
                this.colId = value;
            }
            get
            {
                return this.colId;
            }
        }

        public Int32 RowId
        {
            set
            {
                if (value < 1 || value > XLHelper.MaxRowNumber)
                    throw new ArgumentOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 XLHelper.MaxRowNumber));
                this.rowId = value;
            }
            get
            {
                return this.rowId;
            }
        }

        public Double ColumnOffset
        {
            set
            {
                this.colOffset = value;
            }
            get
            {
                return this.colOffset;
            }
        }

        public Double RowOffset
        {
            set
            {
                this.rowOffset = value;
            }
            get
            {
                return this.rowOffset;
            }
        }

        public Int32 GetZeroBasedColumn()
        {
            return colId - 1;
        }

        public Int32 GetZeroBasedRow()
        {
            return rowId - 1;
        }
    }
}
