using System;

namespace ClosedXML.Excel;

/// <summary>
/// Methods and properties of <see cref="IXLStyle"/>. The <see cref="XLCellFormat"/> has many
/// properties with same name, but different type, so the interface is explicitly implemented.
/// </summary>
internal partial class XLCellFormat : IXLStyle
{
    IXLAlignment IXLStyle.Alignment
    {
        get => Alignment;
        set => Alignment.SetValue(value);
    }

    IXLBorder IXLStyle.Border
    {
        get => Border;
        set => Border.SetValue(value);
    }

    IXLNumberFormat IXLStyle.DateFormat => NumberFormat;

    IXLFill IXLStyle.Fill
    {
        get => Fill;
        set => Fill.SetValue(value);
    }

    IXLFont IXLStyle.Font
    {
        get => Font;
        set => Font.SetFont(value);
    }

    bool IXLStyle.IncludeQuotePrefix
    {
        get => IncludeQuotePrefix;
        set => IncludeQuotePrefix = value;
    }

    IXLNumberFormat IXLStyle.NumberFormat
    {
        get => NumberFormat;
        set => NumberFormat.SetNumberFormat(value.Format);
    }

    IXLProtection IXLStyle.Protection
    {
        get => Protection;
        set => Protection.SetValue(value);
    }

    IXLStyle IXLStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    bool IEquatable<IXLStyle>.Equals(IXLStyle? other)
    {
        if (other is null)
            return false;

        // The API object for each component implement IEquitable<IXL..>
        if (!NumberFormat.Equals(other.NumberFormat))
            return false;

        if (!Font.Equals(other.Font))
            return false;

        if (!IncludeQuotePrefix.Equals(other.IncludeQuotePrefix))
            return false;

        if (!Fill.Equals(other.Fill))
            return false;

        if (!Border.Equals(other.Border))
            return false;

        if (!Alignment.Equals(other.Alignment))
            return false;

        if (!Protection.Equals(other.Protection))
            return false;

        return true;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>rangeApi.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXLStyle value)
    {
        NumberFormat.SetNumberFormat(value.NumberFormat.Format);
        Font.SetFont(value.Font);
        IncludeQuotePrefix = value.IncludeQuotePrefix;
        Fill.SetValue(value.Fill);
        Border.SetValue(value.Border);
        Alignment.SetValue(value.Alignment);
        Protection.SetValue(value.Protection);
    }
}
