using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using System.Linq;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;

namespace ClosedXML.Excel.Charts
{
    public static class ChartGenerator
    {
        public static RadarChart CreateRadarChart(XLChart chartData)
        {
            var radarStyleValue = RadarStyleValues.Marker;
            if (chartData.HasFill)
                radarStyleValue = RadarStyleValues.Filled;
            RadarChart radarChart = new RadarChart();

            RadarStyle radarStyle = new RadarStyle { Val = radarStyleValue };
            radarChart.Append(radarStyle);

            VaryColors varyColors = new VaryColors { Val = false };
            radarChart.Append(varyColors);

            ChartAxis xAxis = null;
            ChartAxis yAxis = null;
            foreach (var series in chartData.Series)
            {
                if (series.SeriesType == ChartSeriesType.Scatter || series.SeriesType == ChartSeriesType.Radar)
                {
                    var scatterChartSeries = GenerateRadarChartSeries(series, chartData.TableReferenced);
                    radarChart.Append(scatterChartSeries);
                    xAxis = series.XAxis;
                    yAxis = series.YAxis;
                }
            }

            AxisId axisId1 = new AxisId { Val = xAxis.Id };
            AxisId axisId2 = new AxisId { Val = yAxis.Id };
            radarChart.Append(axisId1);
            radarChart.Append(axisId2);

            return radarChart;

        }

        public static OpenXmlElement CreatePieChart(XLChart chartData, OpenXmlElement chart)
        {
            VaryColors varyColors = new VaryColors { Val = true };
            chart.Append(varyColors);

            foreach (var series in chartData.Series)
            {
                if (series.SeriesType == ChartSeriesType.Pie)
                {
                    var pieSeries = GeneratePieChartSeries(series, chartData.TableReferenced);
                    chart.Append(pieSeries);
                }
            }

            chart.Append(new FirstSliceAngle() { Val = 0 });

            if (chart is DoughnutChart)
                chart.Append(new HoleSize() { Val = 50 });

            return chart;
        }

        public static ScatterChart CreateScatterChart(XLChart chartData)
        {
            ScatterChart scatterChart = new ScatterChart();

            ScatterStyle scatterStyle = new ScatterStyle { Val = ScatterStyleValues.LineMarker };
            scatterChart.Append(scatterStyle);

            VaryColors varyColors = new VaryColors { Val = false };
            scatterChart.Append(varyColors);

            ChartAxis xAxis = null;
            ChartAxis yAxis = null;
            foreach (var series in chartData.Series)
            {
                if (series.SeriesType == ChartSeriesType.Scatter)
                {
                    var scatterChartSeries = GenerateScatterChartSeries(series, chartData.TableReferenced);
                    scatterChart.Append(scatterChartSeries);
                    xAxis = series.XAxis;
                    yAxis = series.YAxis;
                }
            }

            AxisId axisId1 = new AxisId { Val = xAxis.Id };
            AxisId axisId2 = new AxisId { Val = yAxis.Id };
            scatterChart.Append(axisId1);
            scatterChart.Append(axisId2);

            return scatterChart;
        }

        public static AreaChart CreateAreaChart(XLChart chartData, GroupingValues groupingValue)
        {
            AreaChart areaChart = new AreaChart();

            Grouping grouping = new Grouping { Val = groupingValue };
            areaChart.Append(grouping);

            VaryColors varyColors = new VaryColors { Val = false };
            areaChart.Append(varyColors);

            ChartAxis xAxis = null;
            ChartAxis yAxis = null;
            foreach (var series in chartData.Series)
            {
                if (series.SeriesType == ChartSeriesType.Area)
                {
                    var areaChartSeries = GenerateAreaChartSeries(series, chartData.TableReferenced);
                    areaChart.Append(areaChartSeries);
                    xAxis = series.XAxis;
                    yAxis = series.YAxis;
                }
            }

            AxisId axisId1 = new AxisId { Val = xAxis.Id };
            AxisId axisId2 = new AxisId { Val = yAxis.Id };
            areaChart.Append(axisId1);
            areaChart.Append(axisId2);

            return areaChart;
        }

        public static LineChart CreateLineChart(XLChart chartData)
        {
            LineChart lineChart = new LineChart();

            Grouping grouping = new Grouping() { Val = GroupingValues.Standard };
            lineChart.Append(grouping);

            VaryColors varyColors = new VaryColors { Val = false };
            lineChart.Append(varyColors);

            ChartAxis xAxis = null;
            ChartAxis yAxis = null;
            foreach (var series in chartData.Series)
            {
                if (series.SeriesType == ChartSeriesType.Line)
                {
                    var lineChartSeries = GenerateLineChartSeries(series, chartData.TableReferenced);
                    lineChart.Append(lineChartSeries);
                    xAxis = series.XAxis;
                    yAxis = series.YAxis;
                }
            }

            ShowMarker showMarker = new ShowMarker { Val = chartData.ShowMarkers };
            lineChart.Append(showMarker);

            Smooth smooth = new Smooth { Val = false };
            lineChart.Append(smooth);

            AxisId axisId1 = new AxisId { Val = xAxis.Id };
            AxisId axisId2 = new AxisId { Val = yAxis.Id };
            lineChart.Append(axisId1);
            lineChart.Append(axisId2);

            return lineChart;
        }

        public static BarChart CreateBarChart(XLChart chartData, BarDirectionValues direction, ChartSeriesType type, BarGroupingValues grouping)
        {
            BarChart barChart = new BarChart();

            BarDirection barDirection = new BarDirection { Val = direction };
            barChart.Append(barDirection);

            BarGrouping barGrouping = new BarGrouping { Val = grouping };
            barChart.Append(barGrouping);

            VaryColors varyColors = new VaryColors { Val = false };
            barChart.Append(varyColors);

            ChartAxis xAxis = null;
            ChartAxis yAxis = null;
            foreach (var serie in chartData.Series)
            {
                var barChartSeries = GenerateBarChartSeries(serie, chartData.TableReferenced);
                barChart.Append(barChartSeries);
                xAxis = serie.XAxis;
                yAxis = serie.YAxis;
            }

            if (chartData.SecondaryValueAxisEnabled)
            {
                GapWidth gapWidth = new GapWidth { Val = (ushort)219U };
                barChart.Append(gapWidth);

                Overlap overlap = new Overlap { Val = -27 };
                barChart.Append(overlap);
            }
            else
            {
                GapWidth gapWidth = new GapWidth { Val = (UInt16)150U };
                barChart.Append(gapWidth);

                Overlap overlap = new Overlap { Val = 100 };
                barChart.Append(overlap);
            }

            AxisId axisId1 = new AxisId { Val = xAxis.Id };
            AxisId axisId2 = new AxisId { Val = yAxis.Id };
            barChart.Append(axisId1);
            barChart.Append(axisId2);

            return barChart;
        }

        private static RadarChartSeries GenerateRadarChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            RadarChartSeries radarChartSeries = new RadarChartSeries();

            radarChartSeries.AddIndexAndOrder(seriesData);

            // Series text
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            radarChartSeries.Append(seriesText);

            radarChartSeries.AddNoFill(seriesData);
            //Lines
            if (!seriesData.HasNoFill)
                radarChartSeries.AddColorFill(seriesData);

            //SeriesMarker
            if (seriesData.ShowMarkers)
            {
                radarChartSeries.Append(ChartStyle.AddMarker(seriesData));
            }
            else
            {
                radarChartSeries.Append(ChartStyle.AddNoMarker());
            }

            //DataLabel
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            radarChartSeries.Append(dataLabels);

            //CategoryData
            CategoryAxisData categoryAxisData = tableReferenced ? ChartPartGenerator.GenerateReferencedCategoryAxisData(seriesData) : ChartPartGenerator.GenerateCategoryAxisData(seriesData);
            radarChartSeries.Append(categoryAxisData);

            //ValueData
            Values values = tableReferenced ? ChartPartGenerator.GenerateReferencedValues(seriesData) : ChartPartGenerator.GenerateValues(seriesData);
            radarChartSeries.Append(values);

            return radarChartSeries;
        }

        private static PieChartSeries GeneratePieChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            PieChartSeries pieSeries = new PieChartSeries();
            pieSeries.AddIndexAndOrder(seriesData);

            // Series text
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            pieSeries.Append(seriesText);

            //DataLabel
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            pieSeries.Append(dataLabels);

            // Category data
            CategoryAxisData categoryAxisData = tableReferenced ? ChartPartGenerator.GenerateReferencedCategoryAxisData(seriesData) : ChartPartGenerator.GenerateCategoryAxisData(seriesData);
            pieSeries.Append(categoryAxisData);

            // Value data
            Values values = tableReferenced ? ChartPartGenerator.GenerateReferencedValues(seriesData) : ChartPartGenerator.GenerateValues(seriesData);
            pieSeries.Append(values);


            return pieSeries;
        }

        private static ScatterChartSeries GenerateScatterChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            ScatterChartSeries scatterChartSeries = new ScatterChartSeries();

            scatterChartSeries.AddIndexAndOrder(seriesData);

            // Series text
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            scatterChartSeries.Append(seriesText);

            //Lines
            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();

            Outline outline = new Outline() { Width = 19050 /*25400, CapType = LineCapValues.Round*/ };
            outline.Append(new NoFill());

            chartShapeProperties.Append(outline);
            scatterChartSeries.Append(chartShapeProperties);

            //SeriesMarker
            scatterChartSeries.Append(ChartStyle.AddMarker(seriesData));

            //DataLabels
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            scatterChartSeries.Append(dataLabels);

            //ValueData
            scatterChartSeries.Append(tableReferenced ? ChartPartGenerator.GenerateReferencedXValues(seriesData) : ChartPartGenerator.GenerateXValues(seriesData));
            scatterChartSeries.Append(tableReferenced ? ChartPartGenerator.GenerateReferencedYValues(seriesData) : ChartPartGenerator.GenerateYValues(seriesData));

            //Line Smoothing
            Smooth smooth = new Smooth { Val = seriesData.IsSmoothLine };
            scatterChartSeries.Append(smooth);

            return scatterChartSeries;
        }

        private static AreaChartSeries GenerateAreaChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            AreaChartSeries areaChartSeries = new AreaChartSeries();

            areaChartSeries.AddIndexAndOrder(seriesData);

            // Series text
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            areaChartSeries.Append(seriesText);

            areaChartSeries.AddNoFill(seriesData);
            if (!seriesData.HasNoFill)
                areaChartSeries.AddColorFill(seriesData);

            //DataLabels
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            areaChartSeries.Append(dataLabels);

            // Category data
            CategoryAxisData categoryAxisData = tableReferenced ? ChartPartGenerator.GenerateReferencedCategoryAxisData(seriesData) : ChartPartGenerator.GenerateCategoryAxisData(seriesData);
            areaChartSeries.Append(categoryAxisData);

            // Value data
            Values values = tableReferenced ? ChartPartGenerator.GenerateReferencedValues(seriesData) : ChartPartGenerator.GenerateValues(seriesData);
            areaChartSeries.Append(values);

            return areaChartSeries;
        }

        private static LineChartSeries GenerateLineChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            LineChartSeries lineChartSeries = new LineChartSeries();

            lineChartSeries.AddIndexAndOrder(seriesData);

            //SeriesText
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            lineChartSeries.Append(seriesText);

            lineChartSeries.AddNoFill(seriesData);
            if (!seriesData.HasNoFill)
                lineChartSeries.AddColorFill(seriesData);

            Marker marker = new Marker { Symbol = new Symbol { Val = MarkerStyleValues.None } };
            lineChartSeries.Append(marker);

            //DataLabels
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            lineChartSeries.Append(dataLabels);

            //CategoryData
            CategoryAxisData categoryAxisData = tableReferenced ? ChartPartGenerator.GenerateReferencedCategoryAxisData(seriesData) : ChartPartGenerator.GenerateCategoryAxisData(seriesData);
            lineChartSeries.Append(categoryAxisData);

            //ValueData
            Values values = tableReferenced ? ChartPartGenerator.GenerateReferencedValues(seriesData) : ChartPartGenerator.GenerateValues(seriesData);
            lineChartSeries.Append(values);

            //Line Smoothing
            Smooth smooth = new Smooth { Val = seriesData.IsSmoothLine };
            lineChartSeries.Append(smooth);

            return lineChartSeries;
        }

        private static BarChartSeries GenerateBarChartSeries(ChartSeriesData seriesData, bool tableReferenced)
        {
            BarChartSeries barChartSeries = new BarChartSeries();

            barChartSeries.AddIndexAndOrder(seriesData);

            //Series text
            SeriesText seriesText = ChartPartGenerator.GenerateSeriesText(seriesData);
            barChartSeries.Append(seriesText);

            //TODO
            InvertIfNegative invertIfNegative = new InvertIfNegative { Val = false };
            barChartSeries.Append(invertIfNegative);

            //GenerateDataLabels
            DataLabels dataLabels = ChartPartGenerator.GenerateReferencedDataLables(seriesData);
            barChartSeries.Append(dataLabels);

            //CategoryData
            CategoryAxisData categoryAxisData = tableReferenced ? ChartPartGenerator.GenerateReferencedCategoryAxisData(seriesData) : ChartPartGenerator.GenerateCategoryAxisData(seriesData);
            barChartSeries.Append(categoryAxisData);

            //ValueData
            Values values = tableReferenced ? ChartPartGenerator.GenerateReferencedValues(seriesData) : ChartPartGenerator.GenerateValues(seriesData);
            barChartSeries.Append(values);

            return barChartSeries;
        }

        private static void AddIndexAndOrder(this OpenXmlElement element, ChartSeriesData seriesData)
        {
            Index index = new Index { Val = (uint)seriesData.Index };
            element.Append(index);

            Order order = new Order { Val = (uint)seriesData.Index };
            element.Append(order);
        }

        public static Legend AddChartLegend()
        {
            Legend legend = new Legend();

            LegendPosition legendPosition = new LegendPosition() { Val = LegendPositionValues.Right };
            legend.Append(legendPosition);

            Overlay overlay = new Overlay() { Val = false };
            legend.Append(overlay);

            legend.Append(ChartStyle.SetTextProperties());

            return legend;
        }

        public static void AddChartTitle(DocumentFormat.OpenXml.Drawing.Charts.Chart chart, string title, XLChart chartXl)
        {
            var ctitle = new Title();

            var tx = new ChartText();
            var rich = new RichText();
            var bodyProperties = new BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = TextVerticalOverflowValues.Ellipsis, Vertical = TextVerticalValues.Horizontal, Wrap = TextWrappingValues.Square, Anchor = TextAnchoringTypeValues.Center, AnchorCenter = true };
            var lstSyle = new ListStyle();

            var paragraph = new Paragraph();
            var paraProp = new ParagraphProperties();
            var defRPr = new DefaultRunProperties() { Bold = false, Italic = false, Underline = TextUnderlineValues.None, Strike = TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0, FontSize = 1400 };
            var solidFill = new SolidFill();
            var schemeColor = new SchemeColor() { Val = SchemeColorValues.Text1 };
            var lumMod = new LuminanceModulation() { Val = 65000 };
            var lumOff = new LuminanceOffset() { Val = 35000 };
            schemeColor.Append(lumMod); schemeColor.Append(lumOff);
            solidFill.Append(schemeColor);
            defRPr.Append(solidFill);

            var latin = new LatinFont() { Typeface = "+mn-lt" };
            var ea = new EastAsianFont() { Typeface = "+mn-ea" };
            var cs = new ComplexScriptFont() { Typeface = "+mn-cs" };
            defRPr.Append(latin); defRPr.Append(ea); defRPr.Append(cs);

            paraProp.Append(defRPr);
            paragraph.Append(paraProp);

            var r = new Run();
            var rPr = new RunProperties() { Language = "de-DE" };
            var t = new Text(chartXl.ChartType == XLChartType.Pie ? chartXl.GetSeries(0).SeriesName.Value : title);
            r.Append(rPr); r.Append(t);

            paragraph.Append(r);

            rich.Append(bodyProperties); rich.Append(lstSyle); rich.Append(paragraph);
            tx.Append(rich);
            ctitle.Append(tx);

            var overlay = new Overlay() { Val = false };
            ctitle.Append(overlay);

            ctitle.Append(ChartStyle.SetTextProperties());

            chart.AppendChild(ctitle);
        }

        public static void SetPosition(WorksheetPart worksheetPart, ChartPart chartPart, XLChart chart)
        {
            if (worksheetPart.DrawingsPart.WorksheetDrawing == null)
                worksheetPart.DrawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            var nvps = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Any() ?
                (UInt32Value)worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;

            var width = chart.Size.Width + chart.ChartPosition.X;
            var height = chart.Size.Height + chart.ChartPosition.Y;

            var columnId = chart.ChartPosition.X;
            var rowId = chart.ChartPosition.Y;
            Xdr.FromMarker fMark = new Xdr.FromMarker
            {
                ColumnId = new Xdr.ColumnId(columnId.ToString()),
                RowId = new Xdr.RowId(rowId.ToString()),
                ColumnOffset = new Xdr.ColumnOffset("0"),
                RowOffset = new Xdr.RowOffset("0")
            };
            Xdr.ToMarker tMark = new Xdr.ToMarker
            {
                ColumnId = new Xdr.ColumnId(width.ToString()),
                RowId = new Xdr.RowId(height.ToString()),
                ColumnOffset = new Xdr.ColumnOffset("0"),
                RowOffset = new Xdr.RowOffset("0")
            };

            var twoCellAnchor = new Xdr.TwoCellAnchor(fMark, tMark,
                new Xdr.GraphicFrame
                {
                    Macro = "",
                    NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties() { Id = nvpId, Name = chart.ChartTitle },
                        new Xdr.NonVisualGraphicFrameDrawingProperties()),
                    Transform = new Xdr.Transform(
                        new Offset() { X = 0L, Y = 0L },
                        new Extents() { Cx = 0L, Cy = 0L }),
                    Graphic = new Graphic(
                        new GraphicData(
                            new ChartReference() { Id = worksheetPart.DrawingsPart.GetIdOfPart(chartPart) })
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }),
                },
                new Xdr.ClientData());

            worksheetPart.DrawingsPart.WorksheetDrawing.Append(twoCellAnchor);
        }

        private static void AddNoFill(this OpenXmlElement element, ChartSeriesData seriesData)
        {
            if (seriesData.HasNoFill)
            {
                element.Append(ChartStyle.GetNoFillProperties(false));
            }
        }

        private static void AddColorFill(this OpenXmlElement element, ChartSeriesData seriesData)
        {
            SolidFill solidFill = new SolidFill();

            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();

            if (seriesData.SeriesType == ChartSeriesType.Line || seriesData.SeriesType == ChartSeriesType.Radar)
            {
                if (seriesData.HasLine)
                    chartShapeProperties.Append(new Outline(solidFill));
                else
                    chartShapeProperties.Append(new Outline(new NoFill()));
            }
            else
                chartShapeProperties.Append(solidFill);

            element.Append(chartShapeProperties);
        }
    }
}
