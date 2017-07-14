using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal enum XLChartTypeCategory { Bar3D }
    internal enum XLBarOrientation { Vertical, Horizontal }
    internal enum XLBarGrouping { Clustered, Percent, Stacked, Standard }
    internal class XLChart: XLDrawing<IXLChart>, IXLChart
    {
        internal IXLWorksheet worksheet;
        public XLChart(XLWorksheet worksheet)
        {
            Container = this;
            this.worksheet = worksheet;
            Int32 zOrder;
            if (worksheet.Charts.Any())
                zOrder = worksheet.Charts.Max(c => c.ZOrder) + 1;
            else
                zOrder = 1;
            ZOrder = zOrder;
            ShapeId = worksheet.Workbook.ShapeIdManager.GetNext();
            RightAngleAxes = true;
        }

        public Boolean RightAngleAxes { get; set; }
        public IXLChart SetRightAngleAxes()
        {
            RightAngleAxes = true;
            return this;
        }
        public IXLChart SetRightAngleAxes(Boolean rightAngleAxes)
        {
            RightAngleAxes = rightAngleAxes;
            return this;
        }

        public XLChartType ChartType { get; set; }
        public IXLChart SetChartType(XLChartType chartType)
        {
            ChartType = chartType;
            return this;
        }

        public XLChartTypeCategory ChartTypeCategory
        {
            get
            {
                if (Bar3DCharts.Contains(ChartType))
                    return XLChartTypeCategory.Bar3D;
                else
                    throw new NotImplementedException();

            }
        }

        private HashSet<XLChartType> Bar3DCharts = new HashSet<XLChartType> { 
            XLChartType.BarClustered3D, 
            XLChartType.BarStacked100Percent3D, 
            XLChartType.BarStacked3D, 
            XLChartType.Column3D, 
            XLChartType.ColumnClustered3D, 
            XLChartType.ColumnStacked100Percent3D, 
            XLChartType.ColumnStacked3D
        };

        public XLBarOrientation BarOrientation
        {
            get
            {
                if (HorizontalCharts.Contains(ChartType))
                    return XLBarOrientation.Horizontal;
                else
                    return XLBarOrientation.Vertical;
            }
        }

        private HashSet<XLChartType> HorizontalCharts = new HashSet<XLChartType>{
            XLChartType.BarClustered, 
            XLChartType.BarClustered3D, 
            XLChartType.BarStacked, 
            XLChartType.BarStacked100Percent, 
            XLChartType.BarStacked100Percent3D, 
            XLChartType.BarStacked3D, 
            XLChartType.ConeHorizontalClustered, 
            XLChartType.ConeHorizontalStacked, 
            XLChartType.ConeHorizontalStacked100Percent, 
            XLChartType.CylinderHorizontalClustered, 
            XLChartType.CylinderHorizontalStacked, 
            XLChartType.CylinderHorizontalStacked100Percent, 
            XLChartType.PyramidHorizontalClustered, 
            XLChartType.PyramidHorizontalStacked, 
            XLChartType.PyramidHorizontalStacked100Percent
        };

        public XLBarGrouping BarGrouping
        {
            get
            {
                if (ClusteredCharts.Contains(ChartType))
                    return XLBarGrouping.Clustered;
                else if (PercentCharts.Contains(ChartType))
                    return XLBarGrouping.Percent;
                else if (StackedCharts.Contains(ChartType))
                    return XLBarGrouping.Stacked;
                else
                    return XLBarGrouping.Standard;
            }
        }

        public HashSet<XLChartType> ClusteredCharts = new HashSet<XLChartType>()
        {
            XLChartType.BarClustered,
            XLChartType.BarClustered3D,
            XLChartType.ColumnClustered,
            XLChartType.ColumnClustered3D,
            XLChartType.ConeClustered,
            XLChartType.ConeHorizontalClustered,
            XLChartType.CylinderClustered,
            XLChartType.CylinderHorizontalClustered,
            XLChartType.PyramidClustered,
            XLChartType.PyramidHorizontalClustered
        };

        public HashSet<XLChartType> PercentCharts = new HashSet<XLChartType>() { 
            XLChartType.AreaStacked100Percent,
            XLChartType.AreaStacked100Percent3D,
            XLChartType.BarStacked100Percent,
            XLChartType.BarStacked100Percent3D,
            XLChartType.ColumnStacked100Percent,
            XLChartType.ColumnStacked100Percent3D,
            XLChartType.ConeHorizontalStacked100Percent,
            XLChartType.ConeStacked100Percent,
            XLChartType.CylinderHorizontalStacked100Percent,
            XLChartType.CylinderStacked100Percent,
            XLChartType.LineStacked100Percent,
            XLChartType.LineWithMarkersStacked100Percent,
            XLChartType.PyramidHorizontalStacked100Percent,
            XLChartType.PyramidStacked100Percent
        };

        public HashSet<XLChartType> StackedCharts = new HashSet<XLChartType>()
        {
            XLChartType.AreaStacked,
            XLChartType.AreaStacked3D,
            XLChartType.BarStacked,
            XLChartType.BarStacked3D,
            XLChartType.ColumnStacked,
            XLChartType.ColumnStacked3D,
            XLChartType.ConeHorizontalStacked,
            XLChartType.ConeStacked,
            XLChartType.CylinderHorizontalStacked,
            XLChartType.CylinderStacked,
            XLChartType.LineStacked,
            XLChartType.LineWithMarkersStacked,
            XLChartType.PyramidHorizontalStacked,
            XLChartType.PyramidStacked
        };
    }
}
