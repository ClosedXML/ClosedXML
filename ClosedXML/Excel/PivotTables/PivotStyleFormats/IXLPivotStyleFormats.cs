// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// An API for setting style of parts consisting of <see cref="XLPivotStyleFormatElement"/>, e.g. grand
/// totals. The enumerator enumerates only existing formats, it doesn't add them.
/// </summary>
public interface IXLPivotStyleFormats : IEnumerable<IXLPivotStyleFormat>
{
    /// <summary>
    /// Get styling object for specified <paramref name="element"/>.
    /// </summary>
    /// <param name="element">Which part do we want style for?</param>
    /// <returns>An API to inspect/modify style of the <paramref name="element"/>.</returns>
    /// <exception cref="ArgumentException">When <see cref="XLPivotStyleFormatElement.None"/> is
    ///     passed as an argument.</exception>
    IXLPivotStyleFormat ForElement(XLPivotStyleFormatElement element);
}
