using System;

namespace ClosedXML.Excel;

internal sealed partial class XLNumberCellFormat : IXLNumberFormat
{
    int IXLNumberFormatBase.NumberFormatId
    {
        get => NumberFormatId;
        set => NumberFormatId = value;
    }

    string IXLNumberFormatBase.Format
    {
        get => Format;
        set => Format = value;
    }

    bool IEquatable<IXLNumberFormatBase>.Equals(IXLNumberFormatBase? other)
    {
        if (other is null)
            return false;

        return other.Format == Format;
    }

    IXLStyle IXLNumberFormat.SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _parent;
    }

    IXLStyle IXLNumberFormat.SetFormat(string value)
    {
        Format = value;
        return _parent;
    }
}
