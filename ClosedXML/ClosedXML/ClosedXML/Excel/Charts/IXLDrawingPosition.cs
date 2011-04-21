using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLDrawingAnchor { MoveAndSizeWithCells, MoveWithCells, Absolute}
    public interface IXLDrawingPosition
    {
        XLDrawingAnchor Anchor { get; set; }
        Int32 ZOrder { get; set; }
    }
}
