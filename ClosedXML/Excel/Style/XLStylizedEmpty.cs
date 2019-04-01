using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLStylizedEmpty : XLStylizedBase, IXLStylized
    {
        public XLStylizedEmpty(IXLStyle defaultStyle)
            : base((defaultStyle as XLStyle)?.Value)
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

        protected override IEnumerable<XLStylizedBase> Children
        {
            get { yield break; }
        }
    }
}
