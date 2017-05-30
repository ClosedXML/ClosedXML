using System;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
    {
        IXLStyle SetNumberFormatId(Int32 value);

        IXLStyle SetFormat(String value);
    }
}
