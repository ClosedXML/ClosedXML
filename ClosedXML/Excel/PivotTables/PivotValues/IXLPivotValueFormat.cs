using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
    {
        IXLPivotValue SetNumberFormatId(Int32 value);

        IXLPivotValue SetFormat(String value);
    }
}
