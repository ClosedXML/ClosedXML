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

        internal override UInt32Value GetFieldOffset()
        {
            return UInt32Value.FromUInt32(unchecked((uint)-2));
        }

        internal override IEnumerable<Int32> Match(XLWorkbook.PivotSourceInfo psi, IXLPivotTable pt)
        {
            return new Int32[]
            {
                pt.Values.IndexOf(Value.ToString())
            };
        }
    }
}
