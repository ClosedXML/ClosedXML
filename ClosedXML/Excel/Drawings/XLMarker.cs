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
    }
}
