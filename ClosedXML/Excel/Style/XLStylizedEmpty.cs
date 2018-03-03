using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLStylizedEmpty: XLStylizedBase, IXLStylized
    {
        public XLStylizedEmpty(IXLStyle defaultStyle) : base(defaultStyle?.Value ?? XLStyle.Default.Value)
        {
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
            }
        }

        public override IXLRanges RangesUsed
        {
            get { return new XLRanges(); }
        }
    }
}
