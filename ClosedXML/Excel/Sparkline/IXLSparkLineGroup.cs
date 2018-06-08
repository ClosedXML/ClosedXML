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
        None            = 0,
        HighPoint       = 1 << 1,
        LowPoint        = 1 << 2,
        FirstPoint      = 1 << 3,
        LastPoint       = 1 << 4,
        NegativePoints  = 1 << 5,
        Markers         = 1 << 6,
        All = HighPoint | LowPoint | FirstPoint | LastPoint | NegativePoints | Markers
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

        IXLRange DateRange { get; set; }

        XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        Boolean DisplayHidden { get; set; }

        IXLSparklineHorizontalAxis HorizontalAxis { get; }

        Double LineWeight { get; set; }

        XLSparklineMarkers ShowMarkers { get; set; }

        IXLSparklineStyle Style { get; set; }

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

        IXLSparklineGroup SetDateRange(IXLRange value);

        IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues value);

        IXLSparklineGroup SetDisplayHidden(Boolean value);

        IXLSparklineGroup SetLineWeight(Double value);

        IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value);

        IXLSparklineGroup SetStyle(IXLSparklineStyle value);

        IXLSparklineGroup SetType(XLSparklineType value);

        #endregion Public Methods
    }
}
