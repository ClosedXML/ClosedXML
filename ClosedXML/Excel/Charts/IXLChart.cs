using ClosedXML.Excel.Charts;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLChartType
    {
        Area,
        Area3D,
        AreaStacked,
        AreaStacked100Percent,
        AreaStacked100Percent3D,
        AreaStacked3D,
        BarClustered,
        BarClustered3D,
        BarStacked,
        BarStacked100Percent,
        BarStacked100Percent3D,
        BarStacked3D,
        Bubble,
        Bubble3D,
        Column3D,
        ColumnClustered,
        ColumnClustered3D,
        ColumnStacked,
        ColumnStacked100Percent,
        ColumnStacked100Percent3D,
        ColumnStacked3D,
        Cone,
        ConeClustered,
        ConeHorizontalClustered,
        ConeHorizontalStacked,
        ConeHorizontalStacked100Percent,
        ConeStacked,
        ConeStacked100Percent,
        Cylinder,
        CylinderClustered,
        CylinderHorizontalClustered,
        CylinderHorizontalStacked,
        CylinderHorizontalStacked100Percent,
        CylinderStacked,
        CylinderStacked100Percent,
        Doughnut,
        DoughnutExploded,
        Line,
        Line3D,
        LineStacked,
        LineStacked100Percent,
        LineWithMarkers,
        LineWithMarkersStacked,
        LineWithMarkersStacked100Percent,
        Pie,
        Pie3D,
        PieExploded,
        PieExploded3D,
        PieToBar,
        PieToPie,
        Pyramid,
        PyramidClustered,
        PyramidHorizontalClustered,
        PyramidHorizontalStacked,
        PyramidHorizontalStacked100Percent,
        PyramidStacked,
        PyramidStacked100Percent,
        Radar,
        RadarFilled,
        RadarWithMarkers,
        StockHighLowClose,
        StockOpenHighLowClose,
        StockVolumeHighLowClose,
        StockVolumeOpenHighLowClose,
        Surface,
        SurfaceContour,
        SurfaceContourWireframe,
        SurfaceWireframe,
        XYScatterMarkers,
        XYScatterSmoothLinesNoMarkers,
        XYScatterSmoothLinesWithMarkers,
        XYScatterStraightLinesNoMarkers,
        XYScatterStraightLinesWithMarkers
    }
    public interface IXLChart : IXLDrawing<IXLChart>
    {
        String RelId { get; set; }
        IXLWorksheet Worksheet { get; set; }

        System.Drawing.Point ChartPosition { get; set; }

        XLChartType ChartType { get; set; }
        string ChartTitle { get; set; }
        System.Drawing.Size Size { get; set; }

        bool SecondaryValueAxisEnabled { get; }
        bool ShowLegend { get; set; }
        bool Border { get; set; }
        bool ShowMarkers { get; set; }
        bool TableReferenced { get; set; }

        int SeriesCount { get; }

        List<ChartAxis> Axes { get; set; }

        ChartSeriesData AddSeries(ReferenceData category, ReferenceData value);
        void AddSeries(ChartSeriesData series);

        XLChart CopyPieChart(IXLChart chartToCopy);

        ChartSeriesData GetSeries(int index);

        List<ChartSeriesData> GetAllSeries();

        void DeleteSeries();

    }
}
