using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingWeb
    {
        string AlternateText { get; set; }
        IXLDrawingStyle SetAlternateText(string value);

    }
}
