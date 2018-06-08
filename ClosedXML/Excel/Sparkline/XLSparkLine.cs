using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLSparkline : IXLSparkline
    {        
        public IXLSparklineGroup SparklineGroup { get; }

        public XLFormula Formula { get; set; }
        public IXLCell Cell { get; set; }

        public IXLSparkline SetFormula(XLFormula value) { Formula = value; return this; }
        public IXLSparkline SetCell (IXLCell value) { Cell = value; return this; }

        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
        }

        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup, XLFormula formula)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
            Formula = formula;
        }

        public XLSparkline(IXLCell cell, IXLSparklineGroup sparklineGroup, String formulaText)
        {
            Cell = cell;
            SparklineGroup = sparklineGroup;
            Formula.Value = formulaText;
        }
    }
}
