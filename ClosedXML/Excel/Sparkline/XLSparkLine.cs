using System;

namespace ClosedXML.Excel
{
    internal class XLSparkline : IXLSparkline
    {        
        public IXLSparklineGroup SparklineGroup { get; }

        public XLFormula Formula { get; set; }
        public IXLCell Cell { get; set; }

        public IXLSparkline SetFormula(XLFormula value) { Formula = value; return this; }
        public IXLSparkline SetCell (IXLCell value) { Cell = value; return this; }

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="cell">The cell to place the sparkline in</param>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
        }

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="cell">The cell to place the sparkline in</param>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        /// <param name="formula">The formula for the source range of the sparkline</param>
        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup, XLFormula formula)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
            Formula = formula;
        }

        /// <summary>
        /// Create a new sparkline
        /// </summary>
        /// <param name="cell">The cell to place the sparkline in</param>
        /// <param name="sparklineGroup">The sparkline group to add the sparkline to</param>
        /// <param name="formulaText">The text for the formula for the source range of the sparkline</param>
        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup, String formulaText)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
            Formula.Value = formulaText;
        }

        /// <summary>
        /// Returns the cell this sparkline is used in as an IXLRanges object
        /// </summary>
        /// <returns></returns>
        public IXLRanges GetRanges()
        {
            IXLRanges ranges = new XLRanges();

            if (Cell != null)
            {
                ranges.Add(Cell);
            }

            return ranges;
        }
    }
}
