using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class PivotLabelFieldReference : AbstractPivotFieldReference
    {
        private readonly Predicate<Object> predicate;

        public PivotLabelFieldReference(IXLPivotField pivotField)
            : this(pivotField, null)
        { }

        public PivotLabelFieldReference(IXLPivotField pivotField, Predicate<Object> predicate)
        {
            this.PivotField = pivotField ?? throw new ArgumentNullException(nameof(pivotField));
            this.predicate = predicate;
        }

        public IXLPivotField PivotField { get; set; }

        internal override bool DefaultSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Automatic);
        internal override bool SumSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Sum);
        internal override bool CountSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Count);
        internal override bool CountASubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.CountNumbers);
        internal override bool AverageSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Average);
        internal override bool MaxSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Maximum);
        internal override bool MinSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Minimum);
        internal override bool ApplyProductInSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Product);
        internal override bool ApplyVarianceInSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.Variance);
        internal override bool ApplyVariancePInSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.PopulationVariance);
        internal override bool ApplyStandardDeviationInSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.StandardDeviation);
        internal override bool ApplyStandardDeviationPInSubtotal => PivotField.Subtotals.Contains(XLSubtotalFunction.PopulationStandardDeviation);

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32((uint)PivotField.Offset);
        }

        internal override IEnumerable<Int32> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt)
        {
            var values = pti.Fields[PivotField.SourceName].DistinctValues?.ToList();

            if (predicate == null || values == null)
                return new Int32[] { };

            return values.Select((Value, Index) => new { Value, Index })
                .Where(v => predicate.Invoke(v.Value))
                .Select(v => v.Index)
                .ToList();
        }
    }
}
