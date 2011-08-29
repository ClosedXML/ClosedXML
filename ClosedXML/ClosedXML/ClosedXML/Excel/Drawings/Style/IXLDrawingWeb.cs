using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingWeb
    {
        String AlternativeText { get; set; }
        IXLDrawingStyle SetAlternativeText(String value);

    }
}
