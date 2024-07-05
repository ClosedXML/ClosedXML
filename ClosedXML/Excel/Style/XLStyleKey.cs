using System;

namespace ClosedXML.Excel;

internal readonly record struct XLStyleKey
{
    public required XLAlignmentKey Alignment { get; init; }

    public required XLBorderKey Border { get; init; }

    public required XLFillKey Fill { get; init; }

    public required XLFontKey Font { get; init; }

    public required Boolean IncludeQuotePrefix { get; init; }

    public required XLNumberFormatKey NumberFormat { get; init; }

    public required XLProtectionKey Protection { get; init; }

    public override string ToString()
    {
        return
            this == XLStyle.Default.Key ? "Default" :
                string.Format("Alignment: {0} Border: {1} Fill: {2} Font: {3} IncludeQuotePrefix: {4} NumberFormat: {5} Protection: {6}",
                    Alignment == XLStyle.Default.Key.Alignment ? "Default" : Alignment.ToString(),
                    Border == XLStyle.Default.Key.Border ? "Default" : Border.ToString(),
                    Fill == XLStyle.Default.Key.Fill ? "Default" : Fill.ToString(),
                    Font == XLStyle.Default.Key.Font ? "Default" : Font.ToString(),
                    IncludeQuotePrefix == XLStyle.Default.Key.IncludeQuotePrefix ? "Default" : IncludeQuotePrefix.ToString(),
                    NumberFormat == XLStyle.Default.Key.NumberFormat ? "Default" : NumberFormat.ToString(),
                    Protection == XLStyle.Default.Key.Protection ? "Default" : Protection.ToString());
    }

    public void Deconstruct(
        out XLAlignmentKey alignment,
        out XLBorderKey border,
        out XLFillKey fill,
        out XLFontKey font,
        out Boolean includeQuotePrefix,
        out XLNumberFormatKey numberFormat,
        out XLProtectionKey protection)
    {
        alignment = Alignment;
        border = Border;
        fill = Fill;
        font = Font;
        includeQuotePrefix = IncludeQuotePrefix;
        numberFormat = NumberFormat;
        protection = Protection;
    }
}
