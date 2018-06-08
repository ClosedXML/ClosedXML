// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public interface IXLSparklineStyle
    {
        #region Public Properties

        XLColor FirstMarkerColor { get; set; }

        XLColor HighMarkerColor { get; set; }

        XLColor LastMarkerColor { get; set; }

        XLColor LowMarkerColor { get; set; }

        XLColor MarkersColor { get; set; }

        XLColor NegativeColor { get; set; }

        XLColor SeriesColor { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineStyle SetFirstMarkerColor(XLColor value);

        IXLSparklineStyle SetHighMarkerColor(XLColor value);

        IXLSparklineStyle SetLastMarkerColor(XLColor value);

        IXLSparklineStyle SetLowMarkerColor(XLColor value);

        IXLSparklineStyle SetMarkersColor(XLColor value);

        IXLSparklineStyle SetNegativeColor(XLColor value);

        IXLSparklineStyle SetSeriesColor(XLColor value);

        #endregion Public Methods
    }
}
