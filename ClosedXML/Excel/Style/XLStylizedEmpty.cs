using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStylizedEmpty: IXLStylized
    {
        public XLStylizedEmpty(IXLStyle defaultStyle)
        {
            Style = new XLStyle(this, defaultStyle);
        }
        public IXLStyle Style { get; set; }
        
        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return Style;
                UpdatingStyle = false;
            }
        }

        public bool UpdatingStyle { get; set; }

        public IXLStyle InnerStyle { get; set; }

        public IXLRanges RangesUsed
        {
            get { return new XLRanges(); }
        }

        public bool StyleChanged { get; set; }
    }
}
