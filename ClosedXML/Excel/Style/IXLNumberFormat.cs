using System;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
    {
        IXLStyle SetNumberFormatId(int value);

        IXLStyle SetFormat(string value);
    }
}
