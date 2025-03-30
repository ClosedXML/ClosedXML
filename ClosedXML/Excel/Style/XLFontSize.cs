using System;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel;

/// <summary>
/// Size of a font stored in twips. Storing font size as double causes various problems with
/// equality of font size. Per MS-OI29500: Office converts the points provided to twips, and
/// rounding may occur when writing sz@val back to SpreadsheetML files.
/// </summary>
internal readonly record struct XLFontSize(short Twips)
{
    [return: NotNullIfNotNull(nameof(sizeInPoints))]
    public static XLFontSize? FromPoints(double? sizeInPoints)
    {
        if (sizeInPoints is null)
            return null;

        var twips = Math.Round(sizeInPoints.Value * 20, MidpointRounding.AwayFromZero);
        return new XLFontSize(checked((short)twips));
    }

    /// <summary>
    /// Font size converted to points. Can have rounding issues, so use only when necessary.
    /// </summary>
    public double Points => Twips / 20.0;
}
