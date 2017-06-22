using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingPosition: IXLDrawingPosition
    {
        public Int32 Column { get; set; }
        public IXLDrawingPosition SetColumn(Int32 column) { Column = column; return this; }
        public Double ColumnOffset { get; set; }
        public IXLDrawingPosition SetColumnOffset(Double columnOffset) { ColumnOffset = columnOffset; return this; }
         
        public Int32 Row { get; set; }
        public IXLDrawingPosition SetRow(Int32 row) { Row = row; return this; }
        public Double RowOffset { get; set; }
        public IXLDrawingPosition SetRowOffset(Double rowOffset) { RowOffset = rowOffset; return this; }
    }
}
