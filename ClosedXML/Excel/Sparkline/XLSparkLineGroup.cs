using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroup : IXLSparklineGroup
    {
        public XLSparklineGroup(IXLWorksheet worksheet)
        {
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));

            AxisColor = XLColor.Black;
            SeriesColor = XLColor.FromHtml("FF376092");
            MarkersColor = XLColor.FromHtml("FFD00000");
            HighMarkerColor = XLColor.Black;
            LowMarkerColor = XLColor.Black;
            FirstMarkerColor = XLColor.Black;
            LastMarkerColor = XLColor.Black;
            NegativeColor = XLColor.Black;

            LineWeight = 0.75d;
        }

        public XLSparklineGroup(IXLSparklineGroup sparklineGroup, IXLWorksheet worksheet)
            : this(worksheet)
        {    
            CopyFrom(sparklineGroup);
        }

        private readonly Dictionary<IXLCell, IXLSparkline> _sparklines = new Dictionary<IXLCell, IXLSparkline>();
        public IXLSparkline AddSparkline(IXLCell cell)
        {
            var sparkline = new XLSparkline(cell, this);
            _sparklines.Add(cell, sparkline);
            return sparkline;
        }

        public IXLSparkline AddSparkline(IXLCell cell, String formulaText)
        {
            var sparkline = new XLSparkline(cell, this, new XLFormula(formulaText));
            _sparklines.Add(cell, sparkline);
            return sparkline;
        }

        public IXLSparkline AddSparkline(IXLCell cell, XLFormula formula)
        {
            var sparkline = new XLSparkline(cell, this, formula);
            _sparklines.Add(cell, sparkline);
            return sparkline;
        }

        public IEnumerator<IXLSparkline> GetEnumerator()
        {
            return _sparklines.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(IXLCell cell)
        {
            _sparklines.Remove(cell);
        }

        public void Remove(IXLSparkline sparkline)
        {
            foreach (var sl in _sparklines.Where(kvp => kvp.Value == sparkline).ToList())
            {
                _sparklines.Remove(sl.Key);
            }
        }

        public void RemoveAll()
        {
            _sparklines.Clear();
        }

        public XLColor AxisColor { get; set; }
        public IXLSparklineGroup SetAxisColor(XLColor value) { AxisColor = value; return this; }

        public XLColor FirstMarkerColor { get; set; }
        public IXLSparklineGroup SetFirstMarkerColor(XLColor value) { FirstMarkerColor = value; return this; }

        public XLColor LastMarkerColor { get; set; }
        public IXLSparklineGroup SetLastMarkerColor(XLColor value) { LastMarkerColor = value; return this; }

        public XLColor HighMarkerColor { get; set; }
        public IXLSparklineGroup SetHighMarkerColor(XLColor value) { HighMarkerColor = value; return this; }

        public XLColor LowMarkerColor { get; set; }
        public IXLSparklineGroup SetLowMarkerColor(XLColor value) { LowMarkerColor = value; return this; }

        public XLColor SeriesColor { get; set; }
        public IXLSparklineGroup SetSeriesColor(XLColor value) { SeriesColor = value; return this; }

        public XLColor NegativeColor { get; set; }
        public IXLSparklineGroup SetNegativeColor(XLColor value) { NegativeColor = value; return this; }

        public XLColor MarkersColor { get; set; }
        public IXLSparklineGroup SetMarkersColor(XLColor value) { MarkersColor = value; return this; }              
                
        public Boolean Markers { get; set; }
        public IXLSparklineGroup SetMarkers(Boolean markers)
        {
            Markers = markers;
            return this;
        }

        public Boolean High { get; set; }
        public IXLSparklineGroup SetHigh(Boolean high)
        {
            High = high;
            return this;
        }

        public Boolean Low { get; set; }
        public IXLSparklineGroup SetLow(Boolean low)
        {
            Low = low;
            return this;
        }

        public Boolean First { get; set; }
        public IXLSparklineGroup SetFirst(Boolean first)
        {
            First = first;
            return this;
        }

        public Boolean Last { get; set; }
        public IXLSparklineGroup SetLast(Boolean last)
        {
            Last = last;
            return this;
        }

        public Boolean Negative { get; set; }
        public IXLSparklineGroup SetNegative(Boolean negative)
        {
            Negative = negative;
            return this;
        }

        public Boolean DateAxis { get; set; }
        public IXLSparklineGroup SetDateAxis(Boolean dateAxis)
        {
            DateAxis = dateAxis;
            return this;
        }

        public Boolean DisplayXAxis { get; set; }
        public IXLSparklineGroup SetDisplayXAxis(Boolean displayXAxis)
        {
            DisplayXAxis = displayXAxis;
            return this;
        }

        public Boolean DisplayHidden { get; set; }
        public IXLSparklineGroup SetDisplayHidden(Boolean displayHidden)
        {
            DisplayHidden = displayHidden;
            return this;
        }

        public Boolean RightToLeft { get; set; }
        public IXLSparklineGroup SetRightToLeft(Boolean rightToLeft)
        {
            RightToLeft = rightToLeft;
            return this;
        }           

        public Double? ManualMin { get; set; }
        public IXLSparklineGroup SetManualMin(Double? manualMin)
        {
            ManualMin = manualMin;
            return this;
        }

        public Double? ManualMax { get; set; }
        public IXLSparklineGroup SetManualMax(Double? manualMax)
        {
            ManualMax = manualMax;
            return this;
        }

        public Double LineWeight { get; set; }
        public IXLSparklineGroup SetLineWeight(Double lineWeight)
        {
            LineWeight = lineWeight;
            return this;
        }

        public XLSparklineType Type { get; set; }
        public IXLSparklineGroup SetType(XLSparklineType type)
        {
            Type = type;
            return this;
        }

        public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }
        public IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues displayEmptyCellsAs)
        {
            DisplayEmptyCellsAs = displayEmptyCellsAs;
            return this;
        }

        public XLSparklineAxisMinMax MinAxisType { get; set; }
        public IXLSparklineGroup SetMinAxisType(XLSparklineAxisMinMax minAxisType)
        {
            MinAxisType = minAxisType;
            return this;
        }

        public XLSparklineAxisMinMax MaxAxisType { get; set; }
        public IXLSparklineGroup SetMaxAxisType(XLSparklineAxisMinMax maxAxisType)
        {
            MaxAxisType = maxAxisType;
            return this;
        }

        public IXLWorksheet Worksheet { get; }

        public void CopyTo(IXLWorksheet targetSheet)
        {
            var newSlg = targetSheet.SparklineGroups.AddCopy(this, targetSheet);

            foreach (var sl in this)
            {
                newSlg.AddSparkline(targetSheet.Cell(sl.Cell.Address), targetSheet.Range(sl.Formula.Value).RangeAddress.ToStringRelative(true));
            }
        }

        public void CopyFrom(IXLSparklineGroup sparklineGroup)
        {
            AxisColor = sparklineGroup.AxisColor;
            SeriesColor = sparklineGroup.SeriesColor;
            MarkersColor = sparklineGroup.MarkersColor;
            HighMarkerColor = sparklineGroup.HighMarkerColor;
            LowMarkerColor = sparklineGroup.LowMarkerColor;
            FirstMarkerColor = sparklineGroup.FirstMarkerColor;
            LastMarkerColor = sparklineGroup.LastMarkerColor;
            NegativeColor = sparklineGroup.NegativeColor;

            DateAxis = sparklineGroup.DateAxis;
            Markers = sparklineGroup.Markers;
            High = sparklineGroup.High;
            Low = sparklineGroup.Low;
            First = sparklineGroup.First;
            Last = sparklineGroup.Last;
            Negative = sparklineGroup.Negative;
            DisplayXAxis = sparklineGroup.DisplayXAxis;
            DisplayHidden = sparklineGroup.DisplayHidden;

            ManualMax = sparklineGroup.ManualMax;
            ManualMin = sparklineGroup.ManualMin;
            LineWeight = sparklineGroup.LineWeight;

            MinAxisType = sparklineGroup.MinAxisType;
            MaxAxisType = sparklineGroup.MaxAxisType;

            Type = sparklineGroup.Type;
            DisplayEmptyCellsAs = sparklineGroup.DisplayEmptyCellsAs;
        }        
    }
}
