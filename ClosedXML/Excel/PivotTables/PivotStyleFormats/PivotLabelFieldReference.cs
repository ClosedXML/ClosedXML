using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

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

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32((uint)PivotField.Offset);
        }

        internal override IEnumerable<Int32> Match(XLWorkbook.PivotSourceInfo psi, IXLPivotTable pt)
        {
            var values = psi.Fields[PivotField.SourceName].DistinctValues.ToList();

            if (predicate == null)
                return new Int32[] { };

            return values.Select((Value, Index) => new { Value, Index })
                .Where(v => predicate.Invoke(v.Value))
                .Select(v => v.Index)
                .ToList();
        }
    }
}
