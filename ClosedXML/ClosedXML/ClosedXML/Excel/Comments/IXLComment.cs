using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLComment : IXLFormattedText<IXLComment>, IXLDrawing<IXLComment>
    {
        String Author { get; set; }
        IXLComment SetAuthor(String value);

        IXLRichString AddSignature();
        IXLRichString AddSignature(string username);
        IXLRichString AddNewLine();
    }

}
