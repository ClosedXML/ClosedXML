using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class PivotValueFieldReference : AbstractPivotFieldReference
    {
        public PivotValueFieldReference(string value)
        {
            Value = value;
        }

        public string Value { get; }

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32(unchecked((uint)-2));
        }

        internal override IEnumerable<int> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt)
        {
            return new[]
            {
                pt.Values.IndexOf(Value.ToString())
            };
        }
    }
}