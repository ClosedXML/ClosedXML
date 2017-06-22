using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingPosition
    {
        Int32 Column { get; set; }
        IXLDrawingPosition SetColumn(Int32 column);
        Double ColumnOffset { get; set; }
        IXLDrawingPosition SetColumnOffset(Double columnOffset);

        Int32 Row { get; set; }
        IXLDrawingPosition SetRow(Int32 row);
        Double RowOffset { get; set; }
        IXLDrawingPosition SetRowOffset(Double rowOffset);
    }
}
