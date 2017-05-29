using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Drawings
{
    internal class XLMarker : IXLMarker
    {
        private Int32 colId;
        private Int32 rowId;

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

        public Double ColumnOffset { get; set; }

        public Double RowOffset { get; set; }

    }
}
