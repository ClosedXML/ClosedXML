using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLStylizedEmpty : XLStylizedBase
    {
        public XLStylizedEmpty(IXLStyle? defaultStyle)
            : base((defaultStyle as XLStyle)?.Value)
        {
        }

        public override IEnumerable<IXLRange> RangesUsed => Array.Empty<IXLRange>();

        protected override IEnumerable<XLStylizedBase> Children => Array.Empty<XLStylizedBase>();
    }
}
