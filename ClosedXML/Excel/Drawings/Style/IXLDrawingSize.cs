using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingSize
    {
        Boolean AutomaticSize { get; set; }
        Double Height { get; set; }
        Double Width { get; set; }

        IXLDrawingStyle SetAutomaticSize(); IXLDrawingStyle SetAutomaticSize(Boolean value);
        IXLDrawingStyle SetHeight(Double value);
        IXLDrawingStyle SetWidth(Double value);

    }
}
