using System;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormat: IXLNumberFormatBase, IEquatable<IXLNumberFormat>
    {
        IXLStyle SetNumberFormatId(Int32 value);
        IXLStyle SetFormat(String value);
    }
}
