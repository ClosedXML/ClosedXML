// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLDisplayBlanksAsValues
    {
        Interpolate = 0,
        NotPlotted = 1,
        Zero = 2
    }

    public enum XLSparklineAxisMinMax
    {
        Automatic = 0,
        SameForAll = 1,
        Custom = 2
    }

    [Flags]
    public enum XLSparklineMarkers
    {
        None = 0,
        HighPoint,
        LowPoint,
        FirstPoint,
        LastPoint,
        NegativePoints,
        Markers
    }

    public enum XLSparklineType
    {
        Line = 0,
        Column = 1,
        Stacked = 2
    }

    public interface IXLSparklineGroup : IEnumerable<IXLSparkline>
    {
        #region Public Properties

        XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        Boolean DisplayHidden { get; set; }

        XLColor FirstMarkerColor { get; set; }

        XLColor HighMarkerColor { get; set; }

        IXLSparklineHorizontalAxis HorizontalAxis { get; }

        XLColor LastMarkerColor { get; set; }

        Double LineWeight { get; set; }

        XLColor LowMarkerColor { get; set; }

        XLColor MarkersColor { get; set; }

        XLColor NegativeColor { get; set; }

        XLColor SeriesColor { get; set; }

        XLSparklineMarkers ShowMarkers { get; set; }

        XLSparklineType Type { get; set; }

        IXLSparklineVerticalAxis VerticalAxis { get; }
        IXLWorksheet Worksheet { get; }

        #endregion Public Properties

        #region Public Methods

        IXLSparkline Add(IXLCell location, IXLRange sourceData);

        IEnumerable<IXLSparkline> Add(IXLRange locationRange, IXLRange sourceDataRange);

        IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress);

        void CopyFrom(IXLSparklineGroup sparklineGroup);

        IXLSparklineGroup CopyTo(IXLWorksheet targetSheet);

        IXLSparkline GetSparkline(IXLCell cell);

        IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange);

        void Remove(IXLCell cell);

        void Remove(IXLSparkline sparkline);

        void RemoveAll();

        IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues value);

        IXLSparklineGroup SetDisplayHidden(Boolean value);

        IXLSparklineGroup SetFirstMarkerColor(XLColor value);

        IXLSparklineGroup SetHighMarkerColor(XLColor value);

        IXLSparklineGroup SetLastMarkerColor(XLColor value);

        IXLSparklineGroup SetLineWeight(Double value);

        IXLSparklineGroup SetLowMarkerColor(XLColor value);

        IXLSparklineGroup SetMarkersColor(XLColor value);

        IXLSparklineGroup SetNegativeColor(XLColor value);

        IXLSparklineGroup SetSeriesColor(XLColor value);

        IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value);

        IXLSparklineGroup SetType(XLSparklineType value);

        #endregion Public Methods
    }
}
