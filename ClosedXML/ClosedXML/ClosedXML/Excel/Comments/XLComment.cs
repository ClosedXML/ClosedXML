using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {

        public XLComment(IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Container = this;
        }

        public XLComment(XLFormattedText<IXLComment> defaultComment, IXLFontBase defaultFont)
            : base(defaultComment, defaultFont)
        {
            Container = this;
        }

        public XLComment(String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Container = this;
        }

    }

}
