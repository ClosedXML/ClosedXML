using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingMargins
    {
        Boolean Automatic { get; set; }
        Double Left { get; set; }
        Double Right { get; set; }
        Double Top { get; set; }
        Double Bottom { get; set; }
        Double All { set; }

        IXLDrawingStyle SetAutomatic(); IXLDrawingStyle SetAutomatic(Boolean value);
        IXLDrawingStyle SetLeft(Double value);
        IXLDrawingStyle SetRight(Double value);
        IXLDrawingStyle SetTop(Double value);
        IXLDrawingStyle SetBottom(Double value);
        IXLDrawingStyle SetAll(Double value);

    }
}
