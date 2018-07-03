// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLDisplayBlanksAsValues
    {
        Span = 0,
        Gap = 1,
        Zero = 2
    }

    public enum XLSparklineAxisMinMax
    {
        Individual = 0,
        Group = 1,
        Custom = 2
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

        XLColor AxisColor { get; set; }
        Boolean DateAxis { get; set; }

        XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        Boolean DisplayHidden { get; set; }

        Boolean DisplayXAxis { get; set; }

        Boolean First { get; set; }

        XLColor FirstMarkerColor { get; set; }
        Boolean High { get; set; }

        XLColor HighMarkerColor { get; set; }

        Boolean Last { get; set; }

        XLColor LastMarkerColor { get; set; }
        Double LineWeight { get; set; }

        Boolean Low { get; set; }

        XLColor LowMarkerColor { get; set; }
        Double? ManualMax { get; set; }

        Double? ManualMin { get; set; }

        Boolean Markers { get; set; }

        XLColor MarkersColor { get; set; }

        XLSparklineAxisMinMax MaxAxisType { get; set; }

        XLSparklineAxisMinMax MinAxisType { get; set; }

        Boolean Negative { get; set; }

        XLColor NegativeColor { get; set; }

        Boolean RightToLeft { get; set; }

        XLColor SeriesColor { get; set; }

        XLSparklineType Type { get; set; }

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

        IXLSparklineGroup SetAxisColor(XLColor value);

        IXLSparklineGroup SetDateAxis(Boolean value);

        IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues value);

        IXLSparklineGroup SetDisplayHidden(Boolean value);

        IXLSparklineGroup SetDisplayXAxis(Boolean value);

        IXLSparklineGroup SetFirst(Boolean value);

        IXLSparklineGroup SetFirstMarkerColor(XLColor value);

        IXLSparklineGroup SetHigh(Boolean value);

        IXLSparklineGroup SetHighMarkerColor(XLColor value);

        IXLSparklineGroup SetLast(Boolean value);

        IXLSparklineGroup SetLastMarkerColor(XLColor value);

        IXLSparklineGroup SetLineWeight(Double value);

        IXLSparklineGroup SetLow(Boolean value);

        IXLSparklineGroup SetLowMarkerColor(XLColor value);

        IXLSparklineGroup SetManualMax(Double? value);

        IXLSparklineGroup SetManualMin(Double? value);

        IXLSparklineGroup SetMarkers(Boolean value);

        IXLSparklineGroup SetMarkersColor(XLColor value);

        IXLSparklineGroup SetMaxAxisType(XLSparklineAxisMinMax value);

        IXLSparklineGroup SetMinAxisType(XLSparklineAxisMinMax value);

        IXLSparklineGroup SetNegative(Boolean value);

        IXLSparklineGroup SetNegativeColor(XLColor value);

        IXLSparklineGroup SetSeriesColor(XLColor value);

        IXLSparklineGroup SetType(XLSparklineType value);

        #endregion Public Methods
    }
}
