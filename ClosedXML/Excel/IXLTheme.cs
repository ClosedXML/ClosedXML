namespace ClosedXML.Excel
{
    public interface IXLTheme
    {
        XLColor Background1 { get; set; }
        XLColor Text1 { get; set; }
        XLColor Background2 { get; set; }
        XLColor Text2 { get; set; }
        XLColor Accent1 { get; set; }
        XLColor Accent2 { get; set; }
        XLColor Accent3 { get; set; }
        XLColor Accent4 { get; set; }
        XLColor Accent5 { get; set; }
        XLColor Accent6 { get; set; }
        XLColor Hyperlink { get; set; }
        XLColor FollowedHyperlink { get; set; }

        XLColor ResolveThemeColor(XLThemeColor themeColor);
    }
}
