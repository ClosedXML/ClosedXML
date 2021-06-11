using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;

namespace ClosedXML.Excel.Charts
{
    public static class PlotAreaGenerator
    {
        public static PlotArea GeneratePlotArea(XLChart chart)
        {
            PlotArea plotArea = new PlotArea();

            Layout layout = new Layout();
            plotArea.Append(layout);

            if (!chart.CreateChartPerSeries && chart.ChartType != XLChartType.Pie && chart.ChartType != XLChartType.Doughnut)
            {
                AddChart(chart, plotArea);
                GenerateAxes(chart, plotArea);
            }
            else
            {
                AddChart(chart, plotArea);
            }

            return plotArea;
        }

        private static void AddChart(XLChart chart, PlotArea plotArea)
        {
            var barDirectionValue = BarDirectionValues.Column;
            if (chart.Rotated)
                barDirectionValue = BarDirectionValues.Bar;

            switch (chart.ChartType)
            {
                case XLChartType.ColumnClustered:
                case XLChartType.BarClustered: plotArea.Append(ChartGenerator.CreateBarChart(chart, barDirectionValue, ChartSeriesType.Bar, BarGroupingValues.Clustered)); break;
                case XLChartType.ColumnStacked100Percent:
                case XLChartType.BarStacked100Percent: plotArea.Append(ChartGenerator.CreateBarChart(chart, barDirectionValue, ChartSeriesType.Column100Percent, BarGroupingValues.PercentStacked)); break;
                case XLChartType.ColumnStacked:
                case XLChartType.BarStacked: plotArea.Append(ChartGenerator.CreateBarChart(chart, barDirectionValue, ChartSeriesType.Column, BarGroupingValues.Stacked)); break;
                case XLChartType.Line: plotArea.Append(ChartGenerator.CreateLineChart(chart)); break;
                case XLChartType.Area: plotArea.Append(ChartGenerator.CreateAreaChart(chart, GroupingValues.Standard)); break;
                case XLChartType.AreaStacked100Percent: plotArea.Append(ChartGenerator.CreateAreaChart(chart, GroupingValues.PercentStacked)); break;
                case XLChartType.AreaStacked: plotArea.Append(ChartGenerator.CreateAreaChart(chart, GroupingValues.Stacked)); break;
                case XLChartType.XYScatterMarkers: plotArea.Append(ChartGenerator.CreateScatterChart(chart)); break;
                case XLChartType.Pie:
                    PieChart pieChart = new PieChart();
                    plotArea.Append(ChartGenerator.CreatePieChart(chart, pieChart)); break;
                case XLChartType.Doughnut:
                    DoughnutChart doughnutChart = new DoughnutChart();
                    plotArea.Append(ChartGenerator.CreatePieChart(chart, doughnutChart)); break;
                case XLChartType.Radar: plotArea.Append(ChartGenerator.CreateRadarChart(chart)); break;
            }
        }

        private static void GenerateAxes(XLChart chartData, PlotArea plotArea)
        {
            foreach (var axis in chartData.Axes)
            {
                switch (axis.Type)
                {
                    case ChartAxis.ChartAxisType.Category:
                        CategoryAxis categoryAxis = new CategoryAxis();
                        categoryAxis = GenerateAxis(categoryAxis, axis, chartData) as CategoryAxis;
                        plotArea.Append(categoryAxis);
                        break;
                    case ChartAxis.ChartAxisType.ValueGeneric:
                        ValueAxis valueAxisGen = new ValueAxis();
                        valueAxisGen = GenerateAxis(valueAxisGen, axis, chartData) as ValueAxis;
                        plotArea.Append(valueAxisGen);
                        break;
                    case ChartAxis.ChartAxisType.ValuePercent100WithAllTickmarks:
                        ValueAxis valueAxis = new ValueAxis();
                        valueAxis = GenerateValueAxis100Percent(axis);
                        plotArea.Append(valueAxis);
                        break;
                }
            }
        }

        private static OpenXmlCompositeElement GenerateAxis(OpenXmlCompositeElement axisElement, ChartAxis axis, XLChart chartData)
        {
            AxisId axisId = new AxisId { Val = axis.Id };
            axisElement.Append(axisId);

            Scaling scaling = new Scaling();
            Orientation orientation = new Orientation { Val = axis.InvertOrientation ? OrientationValues.MaxMin : OrientationValues.MinMax };
            scaling.Append(orientation);
            axisElement.Append(scaling);

            Delete delete = new Delete { Val = axis.Invisible };
            AxisPosition axisPosition = new AxisPosition { Val = axis.Position };
            axisElement.Append(delete);
            axisElement.Append(axisPosition);

            if (axisElement is ValueAxis || chartData.ChartType == XLChartType.Radar)
            {
                var majorGridline = new MajorGridlines();
                majorGridline.Append(ChartStyle.SetShapeProperties());
                axisElement.Append(majorGridline);
            }

            NumberingFormat numberingFormat = new NumberingFormat { FormatCode = @"General", SourceLinked = true };
            TickLabelPosition tickLabelPosition = new TickLabelPosition();
            if (chartData.HasTickLabel)
            {
                tickLabelPosition = new TickLabelPosition { Val = TickLabelPositionValues.NextTo };
            }
            else
            {
                tickLabelPosition = new TickLabelPosition { Val = TickLabelPositionValues.None };
            }
            CrossingAxis crossingAxis = new CrossingAxis { Val = axis.RelatedAxis.Id };
            Crosses crosses = new Crosses { Val = CrossesValues.AutoZero };

            AutoLabeled autoLabeled1 = new AutoLabeled { Val = true };
            LabelAlignment labelAlignment1 = new LabelAlignment { Val = LabelAlignmentValues.Center };
            LabelOffset labelOffset1 = new LabelOffset { Val = (ushort)100U };
            NoMultiLevelLabels noMultiLevelLabels1 = new NoMultiLevelLabels { Val = false };

            if (axisElement is ValueAxis)
                axisElement.Append(numberingFormat);

            SetTickMark(axisElement, axis);
            axisElement.Append(tickLabelPosition);
            axisElement.Append(ChartStyle.SetTextProperties(axisElement, chartData));
            axisElement.Append(crossingAxis);
            axisElement.Append(crosses);

            if (axisElement is CategoryAxis)
            {
                axisElement.Append(autoLabeled1);
                axisElement.Append(labelAlignment1);
                axisElement.Append(labelOffset1);
                axisElement.Append(noMultiLevelLabels1);
            }
            if (axisElement is CategoryAxis)
                return axisElement as CategoryAxis;

            return axisElement as ValueAxis;
        }

        private static ValueAxis GenerateValueAxis100Percent(ChartAxis axis)
        {
            ValueAxis valueAxis = new ValueAxis();

            AxisId axisId = new AxisId { Val = axis.Id };
            valueAxis.Append(axisId);

            Scaling scaling = new Scaling();
            Orientation orientation = new Orientation { Val = OrientationValues.MinMax };
            scaling.Append(orientation);
            valueAxis.Append(scaling);

            Delete delete = new Delete { Val = axis.Invisible };
            AxisPosition axisPosition = new AxisPosition { Val = axis.Position };
            valueAxis.Append(delete);
            valueAxis.Append(axisPosition);

            var majorGridline = new MajorGridlines();
            majorGridline.Append(ChartStyle.SetShapeProperties());
            valueAxis.Append(majorGridline);

            NumberingFormat numberingFormat = new NumberingFormat { FormatCode = @"0%", SourceLinked = true };
            TickLabelPosition tickLabelPosition = new TickLabelPosition { Val = TickLabelPositionValues.NextTo };
            CrossingAxis crossingAxis = new CrossingAxis { Val = axis.RelatedAxis.Id };
            Crosses crosses = new Crosses { Val = CrossesValues.AutoZero };

            valueAxis.Append(numberingFormat);
            SetTickMark(valueAxis, axis);
            valueAxis.Append(tickLabelPosition);
            valueAxis.Append(ChartStyle.SetTextProperties());
            valueAxis.Append(crossingAxis);
            valueAxis.Append(crosses);

            return valueAxis;
        }

        private static void SetTickMark(OpenXmlCompositeElement axisElement, ChartAxis axis)
        {
            switch (axis.TickMark)
            {
                case 2:
                    axisElement.Append(new MajorTickMark { Val = TickMarkValues.Inside });
                    axisElement.Append(new MinorTickMark { Val = TickMarkValues.Inside });
                    break;
                case 3:
                    axisElement.Append(new MajorTickMark { Val = TickMarkValues.Outside });
                    axisElement.Append(new MinorTickMark { Val = TickMarkValues.Outside });
                    break;
                case 4:
                    axisElement.Append(new MajorTickMark { Val = TickMarkValues.Cross });
                    axisElement.Append(new MinorTickMark { Val = TickMarkValues.Cross });
                    break;
                default:
                    axisElement.Append(new MajorTickMark { Val = TickMarkValues.None });
                    axisElement.Append(new MinorTickMark { Val = TickMarkValues.None });
                    break;
            }
        }
    }
}
