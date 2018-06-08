// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    public static class XLSparklineTheme
    {
        #region Public Properties

        public static IXLSparklineStyle Default => Dark5;

        #region Dark

        public static IXLSparklineStyle Dark1 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Text1, 0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Text1, 0.249977111117893)
        };

        public static IXLSparklineStyle Dark2 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Text1, 0.34998626667073579),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Background1, -0.249977111117893)
        };

        public static IXLSparklineStyle Dark3 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF323232"),
            NegativeColor = XLColor.FromHtml("FFD00000"),
            MarkersColor = XLColor.FromHtml("FFD00000"),
            HighMarkerColor = XLColor.FromHtml("FFD00000"),
            LowMarkerColor = XLColor.FromHtml("FFD00000"),
            FirstMarkerColor = XLColor.FromHtml("FFD00000"),
            LastMarkerColor = XLColor.FromHtml("FFD00000")
        };

        public static IXLSparklineStyle Dark4 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF000000"),
            NegativeColor = XLColor.FromHtml("FF0070C0"),
            MarkersColor = XLColor.FromHtml("FF0070C0"),
            HighMarkerColor = XLColor.FromHtml("FF0070C0"),
            LowMarkerColor = XLColor.FromHtml("FF0070C0"),
            FirstMarkerColor = XLColor.FromHtml("FF0070C0"),
            LastMarkerColor = XLColor.FromHtml("FF0070C0")
        };

        public static IXLSparklineStyle Dark5 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF376092"),
            NegativeColor = XLColor.FromHtml("FFD00000"),
            MarkersColor = XLColor.FromHtml("FFD00000"),
            HighMarkerColor = XLColor.FromHtml("FFD00000"),
            LowMarkerColor = XLColor.FromHtml("FFD00000"),
            FirstMarkerColor = XLColor.FromHtml("FFD00000"),
            LastMarkerColor = XLColor.FromHtml("FFD00000")
        };

        public static IXLSparklineStyle Dark6 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF0070C0"),
            NegativeColor = XLColor.Black,
            MarkersColor = XLColor.Black,
            HighMarkerColor = XLColor.Black,
            LowMarkerColor = XLColor.Black,
            FirstMarkerColor = XLColor.Black,
            LastMarkerColor = XLColor.Black
        };

        #endregion Dark

        #region Colorful

        public static IXLSparklineStyle Colorful1 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF5F5F5F"),
            NegativeColor = XLColor.FromHtml("FFFFB620"),
            MarkersColor = XLColor.FromHtml("FFD70077"),
            HighMarkerColor = XLColor.FromHtml("FF56BE79"),
            LowMarkerColor = XLColor.FromHtml("FFFF5055"),
            FirstMarkerColor = XLColor.FromHtml("FF5687C2"),
            LastMarkerColor = XLColor.FromHtml("FF359CEB")
        };

        public static IXLSparklineStyle Colorful2 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF5687C2"),
            NegativeColor = XLColor.FromHtml("FFFFB620"),
            MarkersColor = XLColor.FromHtml("FFD70077"),
            HighMarkerColor = XLColor.FromHtml("FF56BE79"),
            LowMarkerColor = XLColor.FromHtml("FFFF5055"),
            FirstMarkerColor = XLColor.FromHtml("FF777777"),
            LastMarkerColor = XLColor.FromHtml("FF359CEB")
        };

        public static IXLSparklineStyle Colorful3 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FFC6EFCE"),
            NegativeColor = XLColor.FromHtml("FFFFC7CE"),
            MarkersColor = XLColor.FromHtml("FF8CADD6"),
            HighMarkerColor = XLColor.FromHtml("FF60D276"),
            LowMarkerColor = XLColor.FromHtml("FFFF5367"),
            FirstMarkerColor = XLColor.FromHtml("FFFFDC47"),
            LastMarkerColor = XLColor.FromHtml("FFFFEB9C")
        };

        public static IXLSparklineStyle Colorful4 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromHtml("FF00B050"),
            NegativeColor = XLColor.FromHtml("FFFF0000"),
            MarkersColor = XLColor.FromHtml("FF0070C0"),
            HighMarkerColor = XLColor.FromHtml("FF00B050"),
            LowMarkerColor = XLColor.FromHtml("FFFF0000"),
            FirstMarkerColor = XLColor.FromHtml("FFFFC000"),
            LastMarkerColor = XLColor.FromHtml("FFFFC000")
        };

        public static IXLSparklineStyle Colorful5 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Text2),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent6),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2)
        };

        public static IXLSparklineStyle Colorful6 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Text1),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent6),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2)
        };

        #endregion Colorful

        #region Accent

        public static IXLSparklineStyle Accent1 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent1),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent2),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent2 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent2),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent3),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent3 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent3),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent4),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent4 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent4),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent5),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent5 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent5),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent6),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent6 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent6),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent1),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893)
        };

        #endregion Accent

        #region Accent Darker 25%

        public static IXLSparklineStyle Accent1Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent2),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent2Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent3),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent3Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent4),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent4Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent5),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent5Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent6),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent6Darker25 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent1),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893)
        };

        #endregion Accent Darker 25%

        #region Accent Darker 50%

        public static IXLSparklineStyle Accent1Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent2),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.39997558519241921)
        };

        public static IXLSparklineStyle Accent2Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent3),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.39997558519241921)
        };

        public static IXLSparklineStyle Accent3Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent4),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, 0.39997558519241921)
        };

        public static IXLSparklineStyle Accent4Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent5),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, 0.39997558519241921)
        };

        public static IXLSparklineStyle Accent5Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent6),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, 0.39997558519241921)
        };

        public static IXLSparklineStyle Accent6Darker50 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.499984740745262),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Accent1),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.499984740745262),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, 0.39997558519241921),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, 0.39997558519241921)
        };

        #endregion Accent Darker 50%

        #region Accent Lighter 40%

        public static IXLSparklineStyle Accent1Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent1, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent2Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent3Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent3, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent3, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent3, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent4Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent4, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent4, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent5Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent5, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent5, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent5, -0.249977111117893)
        };

        public static IXLSparklineStyle Accent6Lighter40 => new XLSparklineStyle
        {
            SeriesColor = XLColor.FromTheme(XLThemeColor.Accent6, 0.39997558519241921),
            NegativeColor = XLColor.FromTheme(XLThemeColor.Background1, -0.499984740745262),
            MarkersColor = XLColor.FromTheme(XLThemeColor.Accent6, 0.79998168889431442),
            HighMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.499984740745262),
            LowMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.499984740745262),
            FirstMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893),
            LastMarkerColor = XLColor.FromTheme(XLThemeColor.Accent6, -0.249977111117893)
        };

        #endregion Accent Lighter 40%

        #endregion Public Properties
    }
}
