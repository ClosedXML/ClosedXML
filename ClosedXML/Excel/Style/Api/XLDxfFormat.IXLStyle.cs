using System;

namespace ClosedXML.Excel;

internal partial class XLDxFormat : IXLStyle
{
    IXLAlignment IXLStyle.Alignment
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLBorder IXLStyle.Border
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLNumberFormat IXLStyle.DateFormat => throw new NotImplementedException();

    IXLFill IXLStyle.Fill
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLFont IXLStyle.Font
    {
        get => Font;
        set => Font.SetValue(value);
    }

    bool IXLStyle.IncludeQuotePrefix
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLNumberFormat IXLStyle.NumberFormat
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLProtection IXLStyle.Protection
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLStyle IXLStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        throw new NotImplementedException();
    }

    bool IEquatable<IXLStyle>.Equals(IXLStyle? other)
    {
        throw new NotSupportedException();
    }
}
