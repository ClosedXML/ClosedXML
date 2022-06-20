using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
    {
        IXLPivotValue SetNumberFormatId(int value);

        IXLPivotValue SetFormat(string value);
    }
}
