using System;

namespace ClosedXML.Excel.Charts
{
    public class ChartSeriesData
    {
        ///// <summary>
        ///// Index der Serie
        ///// </summary>
        public int Index { get; set; }

        public ChartSeriesType SeriesType { get; set; }
        public SingleReferenceData SeriesName { get; set; }
        public ReferenceData Category { get; set; }
        public ReferenceData Values { get; set; }
        public String[] Names
        {
            get
            {
                return Values.Values;
            }
        }

        public ChartAxis XAxis { get; set; }
        public ChartAxis YAxis { get; set; }

        public ChartLabel Label { get; set; }

        public bool HasNoFill { get; set; }
        public String HexColor { get; set; }
        public bool HasLine { get; set; }
        public bool ShowMarkers { get; set; }
        public bool IsTrendLine { get; set; }
        public bool IsSmoothLine { get; set; }

        public ChartSeriesData(int seriesElementCount)
        {
            Category = new ReferenceData { Values = new string[seriesElementCount] };
            Values = new ReferenceData { Values = new string[seriesElementCount] };
            Label = new ChartLabel();
        }

        public ChartSeriesData(ReferenceData category, ReferenceData value)
        {
            Category = category;
            Values = value;
        }
    }
}
