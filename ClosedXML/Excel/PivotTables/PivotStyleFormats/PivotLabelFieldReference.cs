#nullable disable

using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class PivotLabelFieldReference : AbstractPivotFieldReference
    {
        private readonly Predicate<XLCellValue> _predicate;

        public PivotLabelFieldReference(IXLPivotField pivotField)
            : this(pivotField, null)
        { }

        public PivotLabelFieldReference(IXLPivotField pivotField, Predicate<XLCellValue> predicate)
        {
            PivotField = pivotField ?? throw new ArgumentNullException(nameof(pivotField));
            _predicate = predicate;
        }

        public IXLPivotField PivotField { get; set; }

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32((uint)PivotField.Offset);
        }

        internal override IEnumerable<Int32> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt)
        {
            var values = pti.Fields[PivotField.SourceName].DistinctValues.ToList();

            if (_predicate == null)
                return Array.Empty<Int32>();

            var result = new List<Int32>();
            for (var i = 0; i < values.Count; ++i)
            {
                var value = values[i];
                if (_predicate(value))
                    result.Add(i);
            }

            return result;
        }
    }
}
