using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueFormat : IXLNumberFormatBase
    {
        IXLPivotValue SetNumberFormatId(Int32 value);

        IXLPivotValue SetFormat(String value);
    }
}
