using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingWeb
    {
        String AlternateText { get; set; }
        IXLDrawingStyle SetAlternateText(String value);

    }
}
