using SkiaSharp;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        internal XLColorKey Key { get; private set; }

        private XLColor(XLColor defaultColor) : this(defaultColor.Key)
        {
        }

        private XLColor() : this(new XLColorKey())
        {
            HasValue = false;
        }

        private XLColor(SKColor color) : this(new XLColorKey
        {
            Color = color,
            ColorType = XLColorType.Color
        })
        {
        }

        private XLColor(int index) : this(new XLColorKey
        {
            Indexed = index,
            ColorType = XLColorType.Indexed
        })
        {
        }

        private XLColor(XLThemeColor themeColor) : this(new XLColorKey
        {
            ThemeColor = themeColor,
            ColorType = XLColorType.Theme
        })
        {
        }

        private XLColor(XLThemeColor themeColor, double themeTint) : this(new XLColorKey
        {
            ThemeColor = themeColor,
            ThemeTint = themeTint,
            ColorType = XLColorType.Theme
        })
        {
        }

        internal XLColor(XLColorKey key)
        {
            Key = key;
            HasValue = true;
        }
    }
}
