using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class PivotLabelFieldReference : AbstractPivotFieldReference
    {
        private readonly Predicate<object> predicate;

        public PivotLabelFieldReference(IXLPivotField pivotField)
            : this(pivotField, null)
        { }

        public PivotLabelFieldReference(IXLPivotField pivotField, Predicate<object> predicate)
        {
            PivotField = pivotField ?? throw new ArgumentNullException(nameof(pivotField));
            this.predicate = predicate;
        }

        public IXLPivotField PivotField { get; set; }

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32((uint)PivotField.Offset);
        }

        internal override IEnumerable<int> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt)
        {
            var values = pti.Fields[PivotField.SourceName].DistinctValues.ToList();

            if (predicate == null)
                return new int[] { };

            return values.Select((Value, Index) => new { Value, Index })
                .Where(v => predicate.Invoke(v.Value))
                .Select(v => v.Index)
                .ToList();
        }
    }
}
