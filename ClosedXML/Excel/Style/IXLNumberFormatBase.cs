using System;

namespace ClosedXML.Excel;

public interface IXLNumberFormatBase
{
    /// <summary>
    /// <para>
    /// Gets a predefined number format. If number format is predefined, the returned value
    /// is a key in <see cref="XLPredefinedFormat.FormatCodes"/>. If the number format
    /// is a custom format, the returned value is undefined (maybe -1, maybe something else).
    /// </para>
    /// <para>
    /// Sets a predefined number format. Use members of <see cref="XLPredefinedFormat.General"/>,
    /// <see cref="XLPredefinedFormat.Number"/> or <see cref="XLPredefinedFormat.DateTime"/>
    /// casted to <c>int</c>.
    /// <example>
    /// <code>
    ///   cell.NumberFormatId = (int)XLPredefinedFormat.Number.Precision2;
    /// </code>
    /// </example>
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The passed value is not a predefined format.</exception>
    int NumberFormatId { get; set; }

    /// <summary>
    /// Gets or sets number format.
    /// </summary>
    string Format { get; set; }
}
