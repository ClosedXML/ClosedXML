namespace ClosedXML.Excel
{
    internal class XLTheme : IXLTheme
    {
        public XLColor Background1 { get; set; }
        public XLColor Text1 { get; set; }
        public XLColor Background2 { get; set; }
        public XLColor Text2 { get; set; }
        public XLColor Accent1 { get; set; }
        public XLColor Accent2 { get; set; }
        public XLColor Accent3 { get; set; }
        public XLColor Accent4 { get; set; }
        public XLColor Accent5 { get; set; }
        public XLColor Accent6 { get; set; }
        public XLColor Hyperlink { get; set; }
        public XLColor FollowedHyperlink { get; set; }

        public XLColor ResolveThemeColor(XLThemeColor themeColor)
        {
            switch (themeColor)
            {
                case XLThemeColor.Background1:
                    return Background1;

                case XLThemeColor.Text1:
                    return Text1;

                case XLThemeColor.Background2:
                    return Background2;

                case XLThemeColor.Text2:
                    return Text2;

                case XLThemeColor.Accent1:
                    return Accent1;

                case XLThemeColor.Accent2:
                    return Accent2;

                case XLThemeColor.Accent3:
                    return Accent3;

                case XLThemeColor.Accent4:
                    return Accent4;
                    
                case XLThemeColor.Accent5:
                    return Accent5;

                case XLThemeColor.Accent6:
                    return Accent6;

                case XLThemeColor.Hyperlink:
                    return Hyperlink;

                case XLThemeColor.FollowedHyperlink:
                    return FollowedHyperlink;

                default:
                    return null;
            }
        }
    }
}
