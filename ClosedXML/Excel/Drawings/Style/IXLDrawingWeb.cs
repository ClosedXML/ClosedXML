using System;

namespace ClosedXML.Excel
{
    public interface IXLDrawingWeb
    {
        String? AlternateText { get; set; }
        IXLDrawingStyle SetAlternateText(String? value);
    }
}
