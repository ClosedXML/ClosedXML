using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public enum XLColorType { Color, Theme, Indexed }
    public enum XLThemeColor
    {
        Background1,
        Text1,
        Background2,
        Text2,
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Hyperlink,
        FollowedHyperlink
    }
    public interface IXLColor: IEquatable<IXLColor>
    {
        XLColorType ColorType { get; }
        Color Color { get;  }
        Int32 Indexed { get;  }
        XLThemeColor ThemeColor { get;  }
        Double ThemeTint { get;  }
        Boolean HasValue { get; }
    }
}
