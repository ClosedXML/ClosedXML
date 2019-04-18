using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class PivotValueFieldReference : AbstractPivotFieldReference
    {
        public PivotValueFieldReference(String value)
        {
            this.Value = value;
        }

        public String Value { get; }

        internal override bool DefaultSubtotal => false;
        internal override bool SumSubtotal => false;
        internal override bool CountSubtotal => false;
        internal override bool CountASubtotal => false;
        internal override bool AverageSubtotal => false;
        internal override bool MaxSubtotal => false;
        internal override bool MinSubtotal => false;
        internal override bool ApplyProductInSubtotal => false;
        internal override bool ApplyVarianceInSubtotal => false;
        internal override bool ApplyVariancePInSubtotal => false;
        internal override bool ApplyStandardDeviationInSubtotal => false;
        internal override bool ApplyStandardDeviationPInSubtotal => false;

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32(unchecked((uint)-2));
        }

        internal override IEnumerable<Int32> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt)
        {
            return new Int32[]
            {
                pt.Values.IndexOf(Value.ToString())
            };
        }
    }
}
