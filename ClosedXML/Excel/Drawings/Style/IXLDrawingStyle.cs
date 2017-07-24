using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingStyle
    {
        //IXLDrawingFont Font { get; }
        IXLDrawingAlignment Alignment { get; }
        IXLDrawingColorsAndLines ColorsAndLines { get; }
        IXLDrawingSize Size { get; }
        IXLDrawingProtection Protection { get; }
        IXLDrawingProperties Properties { get; }
        IXLDrawingMargins Margins { get; }
        IXLDrawingWeb Web { get; }
    }
}
