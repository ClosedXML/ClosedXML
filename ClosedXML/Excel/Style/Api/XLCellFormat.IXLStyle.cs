using System;

namespace ClosedXML.Excel;

/// <summary>
/// Methods and properties of <see cref="IXLStyle"/>. The <see cref="XLCellFormat"/> has many
/// properties with same name, but different type, so the interface is explicitly implemented.
/// </summary>
internal partial class XLCellFormat : IXLStyle
{
    // TODO Styles: Implement remaining format properties by using IXLStyle contract
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
        set => Font.SetFont(value);
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
        throw new NotImplementedException();
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>rangeApi.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXLStyle value)
    {
        throw new NotImplementedException();
    }
}
