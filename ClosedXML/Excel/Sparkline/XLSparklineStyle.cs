// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    internal class XLSparklineStyle : IXLSparklineStyle
    {
        #region Public Properties

        public XLColor FirstMarkerColor { get; set; }

        public XLColor HighMarkerColor { get; set; }

        public XLColor LastMarkerColor { get; set; }

        public XLColor LowMarkerColor { get; set; }

        public XLColor MarkersColor { get; set; }

        public XLColor NegativeColor { get; set; }

        public XLColor SeriesColor { get; set; }

        #endregion Public Properties

        #region Public Methods

        public IXLSparklineStyle SetFirstMarkerColor(XLColor value)
        {
            FirstMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetHighMarkerColor(XLColor value)
        {
            HighMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetLastMarkerColor(XLColor value)
        {
            LastMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetLowMarkerColor(XLColor value)
        {
            LowMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetMarkersColor(XLColor value)
        {
            MarkersColor = value;
            return this;
        }

        public IXLSparklineStyle SetNegativeColor(XLColor value)
        {
            NegativeColor = value;
            return this;
        }

        public IXLSparklineStyle SetSeriesColor(XLColor value)
        {
            SeriesColor = value;
            return this;
        }

        #endregion Public Methods

        public static void Copy(IXLSparklineStyle from, IXLSparklineStyle to)
        {
            to.FirstMarkerColor = from.FirstMarkerColor;
            to.HighMarkerColor = from.HighMarkerColor;
            to.LastMarkerColor = from.LastMarkerColor;
            to.LowMarkerColor = from.LowMarkerColor;
            to.MarkersColor = from.MarkersColor;
            to.NegativeColor = from.NegativeColor;
            to.SeriesColor = from.SeriesColor;
        }
    }
}
