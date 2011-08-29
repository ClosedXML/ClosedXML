using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLComment : IXLFormattedText<IXLComment>, IXLDrawing<IXLComment>
    {
        String Author { get; set; }
        IXLComment SetAuthor(String value);

        Boolean Visible { get; set; }
        IXLComment SetVisible(); IXLComment SetVisible(Boolean value);

    }

}
