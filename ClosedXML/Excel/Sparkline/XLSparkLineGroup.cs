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

        public IXLRange DateRange
        {
            get => _dateRange;
            set => SetDateRange(value);
        }

        public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        public Boolean DisplayHidden { get; set; }

        public IXLSparklineHorizontalAxis HorizontalAxis { get; }

        public Double LineWeight { get; set; }

        public XLSparklineMarkers ShowMarkers { get; set; }

        private IXLSparklineGroups SparklineGroups => Worksheet.SparklineGroups;

        public IXLSparklineStyle Style
        {
            get => _style;
            set => SetStyle(value);
        }

        public XLSparklineType Type { get; set; }

        public IXLSparklineVerticalAxis VerticalAxis { get; }

        /// <summary>
        /// The worksheet this sparkline group is associated with
        /// </summary>
        public IXLWorksheet Worksheet { get; }

        private IXLRange _dateRange;
        private IXLSparklineStyle _style;

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

            return new XLSparkline(this, location, sourceData);
        }

        public IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress)
        {
            var sourceDataRange = Worksheet.Workbook.Range(sourceDataAddress) ??
                                  Worksheet.Range(sourceDataAddress);
            return Add(Worksheet.Range(locationRangeAddress), sourceDataRange);
        }

        /// <summary>
        /// Copy the details from a specified sparkline group
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to copy from</param>
        public void CopyFrom(IXLSparklineGroup sparklineGroup)
        {
            if (sparklineGroup.DateRange != null)
            {
                DateRange = sparklineGroup.DateRange.Worksheet == sparklineGroup.Worksheet
                    ? Worksheet.Range(sparklineGroup.DateRange.RangeAddress.ToString())
                    : sparklineGroup.DateRange;
            }

            DisplayEmptyCellsAs = sparklineGroup.DisplayEmptyCellsAs;
            DisplayHidden = sparklineGroup.DisplayHidden;
            LineWeight = sparklineGroup.LineWeight;
            ShowMarkers = sparklineGroup.ShowMarkers;
            Type = sparklineGroup.Type;

            XLSparklineStyle.Copy(sparklineGroup.Style, Style);
            XLSparklineHorizontalAxis.Copy(sparklineGroup.HorizontalAxis, HorizontalAxis);
            XLSparklineVerticalAxis.Copy(sparklineGroup.VerticalAxis, VerticalAxis);
        }

        /// <summary>
        /// Copy this sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetSheet">The worksheet to copy this sparkline group to</param>
        public IXLSparklineGroup CopyTo(IXLWorksheet targetSheet)
        {
            if (targetSheet == Worksheet)
                throw new InvalidOperationException(
                    "Cannot copy the sparkline group to the same worksheet it belongs to");

            var copy = targetSheet.SparklineGroups.Add(new XLSparklineGroup(targetSheet, this));
            foreach (var sparkline in _sparklines.Values)
            {
                var location = targetSheet.Cell(((XLAddress)sparkline.Location.Address).WithoutWorksheet());
                var sourceData = sparkline.SourceData.Worksheet == Worksheet
                    ? targetSheet.Range(sparkline.SourceData.RangeAddress.ToString())
                    : sparkline.SourceData;

                copy.Add(location, sourceData);
            }
            return copy;
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
            return _sparklines.TryGetValue(cell, out IXLSparkline sparkline) ? sparkline : null;
        }

        public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange)
        {
            foreach (var key in _sparklines.Keys.Where(searchRange.Contains))
            {
                yield return GetSparkline(key);
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

        public IXLSparklineGroup SetDateRange(IXLRange value)
        {
            if (value != null)
            {
                if (value.RowCount() != 1 && value.ColumnCount() != 1)
                    throw new ArgumentException("The date range must be either one row high or one column wide");
            }

            _dateRange = value;
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

        public IXLSparklineGroup SetLineWeight(Double lineWeight)
        {
            LineWeight = lineWeight;
            return this;
        }

        public IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value)
        {
            ShowMarkers = value;
            return this;
        }

        public IXLSparklineGroup SetStyle(IXLSparklineStyle value)
        {
            _style = value ?? throw new ArgumentNullException(nameof(value));
            return this;
        }

        public IXLSparklineGroup SetType(XLSparklineType type)
        {
            Type = type;
            return this;
        }

        internal IXLSparkline Add(IXLSparkline sparkline)
        {
            if (sparkline.Location.Worksheet != Worksheet)
                throw new ArgumentException("The specified sparkline belongs to the different worksheet");

            SparklineGroups.Remove(sparkline.Location);
            _sparklines[sparkline.Location] = sparkline;
            return sparkline;
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
            HorizontalAxis = new XLSparklineHorizontalAxis(this);
            VerticalAxis = new XLSparklineVerticalAxis(this);
            HorizontalAxis.Color = XLColor.Black;
            Style = XLSparklineTheme.Default;
            LineWeight = 0.75d;
        }

        #endregion Private Constructors
    }
}
