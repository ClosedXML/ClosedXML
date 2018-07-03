// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroup : IXLSparklineGroup
    {
        #region Public Properties

        public XLColor AxisColor { get; set; }

        public Boolean DateAxis { get; set; }

        public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        public Boolean DisplayHidden { get; set; }

        public Boolean DisplayXAxis { get; set; }

        public Boolean First { get; set; }

        public XLColor FirstMarkerColor { get; set; }

        public Boolean High { get; set; }

        public XLColor HighMarkerColor { get; set; }

        public Boolean Last { get; set; }

        public XLColor LastMarkerColor { get; set; }

        public Double LineWeight { get; set; }

        public Boolean Low { get; set; }

        public XLColor LowMarkerColor { get; set; }

        public Double? ManualMax { get; set; }

        public Double? ManualMin { get; set; }

        public Boolean Markers { get; set; }

        public XLColor MarkersColor { get; set; }

        public XLSparklineAxisMinMax MaxAxisType { get; set; }

        public XLSparklineAxisMinMax MinAxisType { get; set; }

        public Boolean Negative { get; set; }

        public XLColor NegativeColor { get; set; }

        public Boolean RightToLeft { get; set; }

        public XLColor SeriesColor { get; set; }

        private IXLSparklineGroups SparklineGroups => Worksheet.SparklineGroups;

        public XLSparklineType Type { get; set; }

        /// <summary>
        /// The worksheet this sparkline group is associated with
        /// </summary>
        public IXLWorksheet Worksheet { get; }

        #endregion Public Properties

        #region Public Constructors

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <param name="copyFrom">The sparkline group to copy from</param>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLWorksheet targetWorksheet, IXLSparklineGroup copyFrom)
            : this(targetWorksheet)
        {
            CopyFrom(copyFrom);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLWorksheet targetWorksheet, string locationAddress, string sourceDataAddress)
            : this(targetWorksheet)
        {
            Add(locationAddress, sourceDataAddress);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLCell location, IXLRange sourceData)
            : this(location.Worksheet)
        {
            Add(location, sourceData);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLRange locationRange, IXLRange sourceDataRange)
            : this(locationRange.Worksheet)
        {
            Add(locationRange, sourceDataRange);
        }

        #endregion Public Constructors

        #region Public Methods

        public IEnumerable<IXLSparkline> Add(IXLRange locationRange, IXLRange sourceDataRange)
        {
            var singleRow = locationRange.RowCount() == 1;
            var singleColumn = locationRange.ColumnCount() == 1;
            var newSparklines = new List<IXLSparkline>();

            if (singleRow && singleColumn)
            {
                if (sourceDataRange.RowCount() != 1 && sourceDataRange.ColumnCount() != 1)
                    throw new ArgumentException("sourceDataRange must have either a single row or a single column");
                newSparklines.Add(Add(locationRange.FirstCell(), sourceDataRange));
            }
            else if (singleRow)
            {
                if (locationRange.ColumnCount() != sourceDataRange.ColumnCount())
                    throw new ArgumentException("locationRange and sourceDataRange must have the same width");
                for (int i = 1; i <= locationRange.ColumnCount(); i++)
                {
                    newSparklines.Add(Add(locationRange.Cell(1, i), sourceDataRange.Column(i).AsRange()));
                }
            }
            else if (singleColumn)
            {
                if (locationRange.RowCount() != sourceDataRange.RowCount())
                    throw new ArgumentException("locationRange and sourceDataRange must have the same height");

                for (int i = 1; i <= locationRange.RowCount(); i++)
                {
                    newSparklines.Add(Add(locationRange.Cell(i, 1), sourceDataRange.Row(i).AsRange()));
                }
            }
            else
                throw new ArgumentException("locationRange must have either a single row or a single column");

            return newSparklines;
        }

        /// <summary>
        /// Add a sparkline to the group.
        /// </summary>
        /// <param name="location">The cell to add sparklines to. If it already contains a sparkline
        /// it will be replaced.</param>
        /// <param name="sourceData">The range the sparkline gets data from</param>
        /// <returns>A newly created sparkline.</returns>
        public IXLSparkline Add(IXLCell location, IXLRange sourceData)
        {
            if (location.Worksheet != Worksheet)
                throw new ArgumentException("The specified sparkline belongs to the different worksheet");

            SparklineGroups.Remove(location);

            var sparkline = new XLSparkline(this, location, sourceData);
            _sparklines.Add(location, sparkline);
            return sparkline;
        }

        public IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress)
        {
            return Add(Worksheet.Range(locationRangeAddress), Worksheet.Range(sourceDataAddress));
        }

        /// <summary>
        /// Copy the details from a specified sparkline group
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to copy from</param>
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

        /// <summary>
        /// Copy this sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetSheet">The worksheet to copy this sparkline group to</param>
        public IXLSparklineGroup CopyTo(IXLWorksheet targetSheet)
        {
            if (targetSheet == Worksheet)
                throw new InvalidOperationException(
                    "Cannot copy the sparkline group to the same worksheet it belong to");

            return targetSheet.SparklineGroups.Add(new XLSparklineGroup(targetSheet, this));
        }

        public IEnumerator<IXLSparkline> GetEnumerator()
        {
            return _sparklines.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLSparkline GetSparkline(IXLCell cell)
        {
            return _sparklines.ContainsKey(cell)
                ? _sparklines[cell]
                : null;
        }

        public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange)
        {
            foreach (var cell in searchRange.CellsUsed())
            {
                yield return GetSparkline(cell);
            }
        }

        /// <summary>
        /// Remove all sparklines in the specified cell from this group
        /// </summary>
        /// <param name="cell">The cell to remove sparklines from</param>
        public void Remove(IXLCell cell)
        {
            if (_sparklines.ContainsKey(cell))
                _sparklines.Remove(cell);
        }

        /// <summary>
        /// Remove the sparkline from this group
        /// </summary>
        /// <param name="sparkline"></param>
        public void Remove(IXLSparkline sparkline)
        {
            Remove(sparkline.Location);
        }

        /// <summary>
        /// Remove all sparklines from this group
        /// </summary>
        public void RemoveAll()
        {
            _sparklines.Clear();
        }

        public IXLSparklineGroup SetAxisColor(XLColor value) { AxisColor = value; return this; }

        public IXLSparklineGroup SetDateAxis(Boolean dateAxis)
        {
            DateAxis = dateAxis;
            return this;
        }

        public IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues displayEmptyCellsAs)
        {
            DisplayEmptyCellsAs = displayEmptyCellsAs;
            return this;
        }

        public IXLSparklineGroup SetDisplayHidden(Boolean displayHidden)
        {
            DisplayHidden = displayHidden;
            return this;
        }

        public IXLSparklineGroup SetDisplayXAxis(Boolean displayXAxis)
        {
            DisplayXAxis = displayXAxis;
            return this;
        }

        public IXLSparklineGroup SetFirst(Boolean first)
        {
            First = first;
            return this;
        }

        public IXLSparklineGroup SetFirstMarkerColor(XLColor value) { FirstMarkerColor = value; return this; }

        public IXLSparklineGroup SetHigh(Boolean high)
        {
            High = high;
            return this;
        }

        public IXLSparklineGroup SetHighMarkerColor(XLColor value) { HighMarkerColor = value; return this; }

        public IXLSparklineGroup SetLast(Boolean last)
        {
            Last = last;
            return this;
        }

        public IXLSparklineGroup SetLastMarkerColor(XLColor value) { LastMarkerColor = value; return this; }

        public IXLSparklineGroup SetLineWeight(Double lineWeight)
        {
            LineWeight = lineWeight;
            return this;
        }

        public IXLSparklineGroup SetLow(Boolean low)
        {
            Low = low;
            return this;
        }

        public IXLSparklineGroup SetLowMarkerColor(XLColor value) { LowMarkerColor = value; return this; }

        public IXLSparklineGroup SetManualMax(Double? manualMax)
        {
            ManualMax = manualMax;
            return this;
        }

        public IXLSparklineGroup SetManualMin(Double? manualMin)
        {
            ManualMin = manualMin;
            return this;
        }

        public IXLSparklineGroup SetMarkers(Boolean markers)
        {
            Markers = markers;
            return this;
        }

        public IXLSparklineGroup SetMarkersColor(XLColor value) { MarkersColor = value; return this; }

        public IXLSparklineGroup SetMaxAxisType(XLSparklineAxisMinMax maxAxisType)
        {
            MaxAxisType = maxAxisType;
            return this;
        }

        public IXLSparklineGroup SetMinAxisType(XLSparklineAxisMinMax minAxisType)
        {
            MinAxisType = minAxisType;
            return this;
        }

        public IXLSparklineGroup SetNegative(Boolean negative)
        {
            Negative = negative;
            return this;
        }

        public IXLSparklineGroup SetNegativeColor(XLColor value) { NegativeColor = value; return this; }

        public IXLSparklineGroup SetRightToLeft(Boolean rightToLeft)
        {
            RightToLeft = rightToLeft;
            return this;
        }

        public IXLSparklineGroup SetSeriesColor(XLColor value) { SeriesColor = value; return this; }

        public IXLSparklineGroup SetType(XLSparklineType type)
        {
            Type = type;
            return this;
        }

        #endregion Public Methods

        #region Private Fields

        private readonly Dictionary<IXLCell, IXLSparkline> _sparklines = new Dictionary<IXLCell, IXLSparkline>();

        #endregion Private Fields

        #region Private Constructors

        /// <summary>
        /// Add a new sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <returns>The new sparkline group added</returns>
        internal XLSparklineGroup(IXLWorksheet targetWorksheet)
        {
            Worksheet = targetWorksheet ?? throw new ArgumentNullException(nameof(targetWorksheet));

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

        #endregion Private Constructors
    }
}
