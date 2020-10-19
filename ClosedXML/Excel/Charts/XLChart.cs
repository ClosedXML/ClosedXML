using ClosedXML.Excel.Charts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    public class XLChart : XLDrawing<IXLChart>, IXLChart
    {
        private List<ChartSeriesData> m_series;

        public XLChart(XLChartType type, IXLWorksheet worksheet)
        {
            Container = this;
            this.Worksheet = worksheet;
            ChartType = type;

            m_series = new List<ChartSeriesData>();
            Axes = new List<ChartAxis>();
        }

        public XLChart(XLChartType type)
        {
            ChartType = type;

            m_series = new List<ChartSeriesData>();
            Axes = new List<ChartAxis>();
        }
        public String RelId { get; set; }
        public IXLWorksheet Worksheet { get; set; }
        public System.Drawing.Point ChartPosition { get; set; }

        public XLChartType ChartType { get; set; }
        public string ChartTitle { get; set; }
        public System.Drawing.Size Size { get; set; }
        public int SeriesCount
        {
            get
            {
                return m_series.Count();
            }
        }
        public bool SecondaryValueAxisEnabled
        {
            get
            {
                if (Series.Count() > 1 && ChartType != XLChartType.BarStacked && ChartType != XLChartType.BarStacked100Percent && ChartType != XLChartType.ColumnStacked && ChartType != XLChartType.ColumnStacked100Percent)
                    return true;
                return false;
            }
        }

        public bool ShowLegend { get; set; }
        public bool Border { get; set; }
        public bool ShowMarkers { get; set; }
        public bool CreateChartPerSeries
        {
            get
            {
                if (ChartType == XLChartType.Pie)
                    return true;
                return false;
            }
        }
        public bool Rotated
        {
            get
            {
                if (ChartType != XLChartType.Pie)
                {
                    if (m_series.Any(x => x.SeriesType == ChartSeriesType.Bar))
                        return true;
                }
                return false;
            }
        }
        public bool HasFill { get; set; }
        public bool HasTickLabel { get; set; }
        public bool TableReferenced { get; set; }

        public List<ChartAxis> Axes { get; set; }
        public IEnumerable<ChartSeriesData> Series
        {
            get { return m_series; }
        }

        public ChartSeriesData AddSeries(ReferenceData category, ReferenceData value)
        {
            ChartSeriesData series = new ChartSeriesData(category, value);
            m_series.Add(series);
            return series;
        }
        public void AddSeries(ChartSeriesData series)
        {
            m_series.Add(series);
        }
        public ChartSeriesData GetSeries(int index)
        {
            if (index >= 0 && index < m_series.Count)
                return m_series[index] as ChartSeriesData;
            return null;
        }
        public List<ChartSeriesData> GetAllSeries()
        {
            List<ChartSeriesData> series = new List<ChartSeriesData>();
            foreach (var serie in m_series)
            {
                series.Add(serie);
            }
            return series;
        }
        public void DeleteSeries()
        {
            m_series.Clear();
        }

        public XLChart CopyPieChart(IXLChart chartToCopy)
        {
            int x = chartToCopy.ChartPosition.X;
            int y = chartToCopy.ChartPosition.Y + chartToCopy.Size.Height + 2;
            XLChart newChart = new XLChart(chartToCopy.ChartType)
            {
                Axes = chartToCopy.Axes,
                Border = chartToCopy.Border,
                ChartTitle = chartToCopy.ChartTitle,
                m_series = new List<ChartSeriesData>(),
                ShowLegend = chartToCopy.ShowLegend,
                Size = chartToCopy.Size,
                ChartPosition = new System.Drawing.Point(x, y)
            };
            return newChart;
        }
    }
}
