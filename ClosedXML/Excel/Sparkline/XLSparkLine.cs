// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSparkline : IXLSparkline
    {
        #region Public Properties

        public IXLCell Location { get; set; }

        public IXLRange SourceData { get; set; }

        public IXLSparklineGroup SparklineGroup { get; }

        #endregion Public Properties

        #region Public Constructors

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        /// <param name="cell">The cell to place the sparkline in</param>
        /// <param name="sourceData">The range the sparkline gets data from</param>
        public XLSparkline(IXLSparklineGroup sparklineGroup, IXLCell cell, IXLRange sourceData)
        {
            if (sparklineGroup.Worksheet != cell.Worksheet)
                throw new InvalidOperationException("Cell must belong to the same worksheet as the sparkline group");

            SparklineGroup = sparklineGroup;
            Location = cell;
            SourceData = sourceData;
        }

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        /// <param name="cellAddress">The address of the cell to place the sparkline in</param>
        /// <param name="sourceDataAddress">The address of the sparkline's source range</param>
        public XLSparkline(IXLSparklineGroup sparklineGroup, string cellAddress, string sourceDataAddress)
            : this(sparklineGroup, sparklineGroup.Worksheet.Cell(cellAddress), sparklineGroup.Worksheet.Range(sourceDataAddress))
        {
        }

        #endregion Public Constructors

        #region Public Methods

        public IXLSparkline SetLocation(IXLCell value) { Location = value; return this; }

        public IXLSparkline SetSourceData(IXLRange value) { SourceData = value; return this; }

        #endregion Public Methods
    }
}
