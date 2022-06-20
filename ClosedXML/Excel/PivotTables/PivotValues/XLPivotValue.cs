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
        public string SourceName { get; private set; }
        public string CustomName { get; set; }		public IXLPivotValue SetCustomName(string value) { CustomName = value; return this; }

        public XLPivotSummary SummaryFormula { get; set; }		public IXLPivotValue SetSummaryFormula(XLPivotSummary value) { SummaryFormula = value; return this; }
        public XLPivotCalculation Calculation { get; set; }		public IXLPivotValue SetCalculation(XLPivotCalculation value) { Calculation = value; return this; }
        public string BaseField { get; set; }		public IXLPivotValue SetBaseField(string value) { BaseField = value; return this; }
        public string BaseItem { get; set; }		public IXLPivotValue SetBaseItem(string value) { BaseItem = value; return this; }
        public XLPivotCalculationItem CalculationItem { get; set; }		public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value) { CalculationItem = value; return this; }


        public IXLPivotValue ShowAsNormal()
        {
            return SetCalculation(XLPivotCalculation.Normal);
        }
        public IXLPivotValueCombination ShowAsDifferenceFrom(string fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.DifferenceFrom);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValueCombination ShowAsPercentageFrom(string fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PercentageOf);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValueCombination ShowAsPercentageDifferenceFrom(string fieldSourceName)
        {
            BaseField = fieldSourceName;
            SetCalculation(XLPivotCalculation.PercentageDifferenceFrom);
            return new XLPivotValueCombination(this);
        }
        public IXLPivotValue ShowAsRunningTotalIn(string fieldSourceName)
        {
            BaseField = fieldSourceName;
            return SetCalculation(XLPivotCalculation.RunningTotal);
        }
        public IXLPivotValue ShowAsPercentageOfRow()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfRow);
        }

        public IXLPivotValue ShowAsPercentageOfColumn()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfColumn);
        }

        public IXLPivotValue ShowAsPercentageOfTotal()
        {
            return SetCalculation(XLPivotCalculation.PercentageOfTotal);
        }

        public IXLPivotValue ShowAsIndex()
        {
            return SetCalculation(XLPivotCalculation.Index);
        }
    }
}
