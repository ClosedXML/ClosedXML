#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSparkline : IXLSparkline
    {
        public XLSparklineGroup _sparklineGroup;
        private IXLCell _location;
        private IXLRange _sourceData;

        public bool IsValid =>
            Location != null &&
            SourceData != null &&
            ((XLAddress)Location.Address).IsValid &&
            ((XLRangeAddress)SourceData.RangeAddress).IsValid;

        public IXLCell Location
        {
            get => _location;
            set => SetLocation(value);
        }

        public IXLRange SourceData
        {
            get => _sourceData;
            set => SetSourceData(value);
        }

        public IXLSparklineGroup SparklineGroup => _sparklineGroup;

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        /// <param name="cell">The cell to place the sparkline in</param>
        /// <param name="sourceData">The range the sparkline gets data from</param>
        internal XLSparkline(XLSparklineGroup sparklineGroup, IXLCell cell, IXLRange sourceData)
        {
            if (sparklineGroup == null)
                throw new ArgumentNullException(nameof(sparklineGroup));

            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            if (sparklineGroup.Worksheet != cell.Worksheet)
                throw new InvalidOperationException("Cell must belong to the same worksheet as the sparkline group");

            _sparklineGroup = sparklineGroup;
            Location = cell;
            SourceData = sourceData;
        }

        public IXLSparkline SetLocation(IXLCell cell)
        {
            if (cell != null && cell.Worksheet != SparklineGroup.Worksheet)
                throw new InvalidOperationException("Cannot move the sparkline to a different worksheet");

            if (_location != null)
                SparklineGroup.Remove(_location);

            _location = cell;
            _sparklineGroup.Add(this);
            return this;
        }

        public IXLSparkline SetSourceData(IXLRange range)
        {
            if (range != null && range.RowCount() != 1 && range.ColumnCount() != 1)
                throw new ArgumentException("SourceData range must have either a single row or a single column");

            _sourceData = range;
            return this;
        }
    }
}
