﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotValue: IXLPivotValue
    {
        public XLPivotValue(string sourceName)
        {
            SourceName = sourceName;
            NumberFormat = new XLPivotValueFormat(this);
        }

        public IXLPivotValueFormat NumberFormat { get; private set; }
        public String SourceName { get; private set; }	
        public String CustomName { get; set; }		public IXLPivotValue SetCustomName(String value) { CustomName = value; return this; }

        public XLPivotSummary SummaryFormula { get; set; }		public IXLPivotValue SetSummaryFormula(XLPivotSummary value) { SummaryFormula = value; return this; }
        public XLPivotCalculation Calculation { get; set; }		public IXLPivotValue SetCalculation(XLPivotCalculation value) { Calculation = value; return this; }
        public String BaseField { get; set; }		public IXLPivotValue SetBaseField(String value) { BaseField = value; return this; }
        public String BaseItem { get; set; }		public IXLPivotValue SetBaseItem(String value) { BaseItem = value; return this; }
        public XLPivotCalculationItem CalculationItem { get; set; }		public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value) { CalculationItem = value; return this; }


        public IXLPivotValue ShowAsNormal()
        {
            return SetCalculation(XLPivotCalculation.Normal);
        }
        public IXLPivotValueCombination ShowAsDifferenceFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.DifferenceFrom);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValueCombination ShowAsPctFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PctOf);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValueCombination ShowAsPctDifferenceFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PctDifferenceFrom);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValue ShowAsRunningTotalIn(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            return SetCalculation(XLPivotCalculation.RunningTotal);
        }
        public IXLPivotValue ShowAsPctOfRow()
        {
            return SetCalculation(XLPivotCalculation.PctOfRow);
        }

        public IXLPivotValue ShowAsPctOfColumn()
        {
            return SetCalculation(XLPivotCalculation.PctOfColumn);
        }

        public IXLPivotValue ShowAsPctOfTotal()
        {
            return SetCalculation(XLPivotCalculation.PctOfTotal);
        }

        public IXLPivotValue ShowAsIndex()
        {
            return SetCalculation(XLPivotCalculation.Index);
        }
    }
}
