using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        internal Boolean HasValue { get; private set; }
        internal XLColor(XLColor defaultColor)
        {
            if (defaultColor.ColorType == XLColorType.Theme)
            {
                this.ThemeColor = defaultColor.ThemeColor;
                this.ThemeTint = defaultColor.ThemeTint;
            }
            else if (defaultColor.ColorType == XLColorType.Indexed)
            {
                this.Indexed = defaultColor.Indexed;
            }
            else
            {
                this.Color = defaultColor.Color;
            }
        }
    }
}
