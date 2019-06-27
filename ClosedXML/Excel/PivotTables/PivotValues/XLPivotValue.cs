// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotValue : IXLPivotValue
    {
        public XLPivotValue(string sourceName)
        {
            SourceName = sourceName;
            NumberFormat = new XLPivotValueFormat(this);
        }

        public String BaseField { get; set; }
        public String BaseItem { get; set; }
        public XLPivotCalculation Calculation { get; set; }
        public XLPivotCalculationItem CalculationItem { get; set; }
        public String CustomName { get; set; }
        public IXLPivotValueFormat NumberFormat { get; private set; }
        public String SourceName { get; private set; }
        public XLPivotSummary SummaryFormula { get; set; }

        public IXLPivotValue SetBaseField(String value)
        {
            BaseField = value; return this;
        }

        public IXLPivotValue SetBaseItem(String value)
        {
            BaseItem = value; return this;
        }

        public IXLPivotValue SetCalculation(XLPivotCalculation value)
        {
            Calculation = value; return this;
        }

        public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value)
        {
            CalculationItem = value; return this;
        }

        public IXLPivotValue SetCustomName(String value)
        {
            CustomName = value; return this;
        }

        public IXLPivotValue SetSummaryFormula(XLPivotSummary value)
        {
            SummaryFormula = value; return this;
        }

        public IXLPivotValueCombination ShowAsDifferenceFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.DifferenceFrom);
            return new XLPivotValueCombination(this);
        }

        public IXLPivotValue ShowAsIndex()
        {
            return SetCalculation(XLPivotCalculation.Index);
        }

        public IXLPivotValue ShowAsNormal()
        {
            return SetCalculation(XLPivotCalculation.Normal);
        }

        public IXLPivotValueCombination ShowAsPercentageDifferenceFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PercentageDifferenceFrom);
            return new XLPivotValueCombination(this);
        }

        public IXLPivotValueCombination ShowAsPercentageFrom(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PercentageOf);
            return new XLPivotValueCombination(this);
        }

        public IXLPivotValue ShowAsPercentageOfColumn()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfColumn);
        }

        public IXLPivotValue ShowAsPercentageOfRow()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfRow);
        }

        public IXLPivotValue ShowAsPercentageOfTotal()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfTotal);
        }

        public IXLPivotValue ShowAsRunningTotalIn(String fieldSourceName)
        {
            BaseField = fieldSourceName;
            return SetCalculation(XLPivotCalculation.RunningTotal);
        }
    }
}
