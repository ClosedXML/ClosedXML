using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLSparklineType
    {
        Line = 0,
        Column = 1,
        Stacked = 2        
    }

    public enum XLSparklineAxisMinMax
    {
        Individual = 0,
        Group = 1,
        Custom = 2
    }

    public enum XLDisplayBlanksAsValues
    {
        Span = 0,
        Gap = 1,
        Zero = 2        
    }

    public interface IXLSparklineGroup : IEnumerable<IXLSparkline>
    {
        XLColor AxisColor { get; set; }
        XLColor FirstMarkerColor { get; set; }
        XLColor LastMarkerColor { get; set; }
        XLColor HighMarkerColor { get; set; }
        XLColor LowMarkerColor { get; set; }
        XLColor SeriesColor { get; set; }
        XLColor NegativeColor { get; set; }
        XLColor MarkersColor { get; set; }

        Boolean Markers { get; set; }
        Boolean High { get; set; }
        Boolean Low { get; set; }
        Boolean First { get; set; }
        Boolean Last { get; set; }
        Boolean Negative { get; set; }
        Boolean DateAxis { get; set; }
        Boolean DisplayXAxis { get; set; }
        Boolean DisplayHidden { get; set; }
        Boolean RightToLeft { get; set; }
        Double ManualMax { get; set; }
        Double ManualMin { get; set; }
        Double LineWeight { get; set; }

        XLSparklineAxisMinMax MinAxisType { get; set; }
        XLSparklineAxisMinMax MaxAxisType { get; set; }        

        XLSparklineType Type { get; set; }
        XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }
        
        IXLSparklineGroup SetAxisColor(XLColor value);
        IXLSparklineGroup SetSeriesColor(XLColor value);
        IXLSparklineGroup SetNegativeColor(XLColor value);
        IXLSparklineGroup SetMarkersColor(XLColor value);
        IXLSparklineGroup SetFirstMarkerColor(XLColor value);
        IXLSparklineGroup SetLastMarkerColor(XLColor value);
        IXLSparklineGroup SetHighMarkerColor(XLColor value);
        IXLSparklineGroup SetLowMarkerColor(XLColor value);
        
        IXLSparklineGroup SetDateAxis(Boolean value);
        IXLSparklineGroup SetMarkers(Boolean value);
        IXLSparklineGroup SetHigh(Boolean value);
        IXLSparklineGroup SetLow(Boolean value);
        IXLSparklineGroup SetFirst(Boolean value);
        IXLSparklineGroup SetLast(Boolean value);
        IXLSparklineGroup SetNegative(Boolean value);
        IXLSparklineGroup SetDisplayXAxis(Boolean value);
        IXLSparklineGroup SetDisplayHidden(Boolean value);

        IXLSparklineGroup SetManualMax(Double value);
        IXLSparklineGroup SetManualMin(Double value);
        IXLSparklineGroup SetLineWeight(Double value);

        IXLSparklineGroup SetMinAxisType(XLSparklineAxisMinMax value);
        IXLSparklineGroup SetMaxAxisType(XLSparklineAxisMinMax value);

        IXLSparklineGroup SetType(XLSparklineType value);
        IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues value);

        IXLSparkline AddSparkline(IXLCell cell);
        IXLSparkline AddSparkline(IXLCell cell, string formulaText);
        IXLSparkline AddSparkline(IXLCell cell, XLFormula formula);

        void RemoveAll();

        void Remove(IXLCell cell);

        void Remove(IXLSparkline sparkline);

        void CopyTo(IXLWorksheet targetSheet);

        void CopyFrom(IXLSparklineGroup sparklineGroup);

        IXLWorksheet Worksheet { get; }
    }
}
