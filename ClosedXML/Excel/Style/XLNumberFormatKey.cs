using System;

namespace ClosedXML.Excel;

internal readonly record struct XLNumberFormatKey
{
    /// <summary>
    /// The value <c>-1</c> that is set to <see cref="NumberFormatId"/>, if <see cref="Format"/> is
    /// set to user-defined format (non-empty string).
    /// </summary>
    public const int CustomFormatNumberId = -1;

    /// <summary>
    /// Number format identifier of predefined format, see <see cref="XLPredefinedFormat"/>.
    /// If -1, the format is custom and stored in the <see cref="Format"/>.
    /// </summary>
    public required int NumberFormatId { get; init; }

    public required string Format { get; init; }

    public static XLNumberFormatKey ForFormat(string customFormat)
    {
        if (string.IsNullOrEmpty(customFormat))
            throw new ArgumentException();

        return new XLNumberFormatKey
        {
            NumberFormatId = CustomFormatNumberId,
            Format = customFormat,
        };
    }

    public override string ToString()
    {
        return $"{Format}/{NumberFormatId}";
    }
}
