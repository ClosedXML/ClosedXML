using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStylizedContainer: IXLStylized
    {
        IXLStylized container;
        public XLStylizedContainer(IXLStyle style, IXLStylized container)
        {
            this.Style = style;
            this.container = container;
        }

        public IXLStyle Style { get; set; }

        public IEnumerable<IXLStyle> Styles
        {
            get 
            {
                container.UpdatingStyle = true;
                yield return Style;
                container.UpdatingStyle = false;
            }
        }

        public bool UpdatingStyle { get; set; }

        public IXLStyle InnerStyle { get; set; }
    }
}
