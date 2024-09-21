using System.Collections.Generic;

namespace ClosedXML.Excel;

public interface IXLHyperlinks: IEnumerable<XLHyperlink>
{
    /// <summary>
    /// Remove the hyperlink from a worksheet. Doesn't throw if hyperlinks is
    /// not attached to a worksheet.
    /// </summary>
    /// <remarks>
    /// If hyperlink range uses a <see cref="XLThemeColor.Hyperlink">hyperlink
    /// theme color</see>, the style is reset to the sheet style font color.
    /// The <see cref="IXLFontBase.Underline"/> is also set to sheet style
    /// underline.
    /// </remarks>
    /// <param name="hyperlink">Hyperlink to remove.</param>
    /// <returns><c>true</c> if hyperlink was part of the worksheet and was
    ///   removed. <c>false</c> otherwise.</returns>
    bool Delete(XLHyperlink hyperlink);

    /// <summary>
    /// Delete a hyperlink defined for a single cell. It doesn't delete
    /// hyperlinks that cover the cell.
    /// </summary>
    /// <remarks>
    /// If hyperlink range uses a <see cref="XLThemeColor.Hyperlink">hyperlink
    /// theme color</see>, the style is reset to the sheet style font color.
    /// The <see cref="IXLFontBase.Underline"/> is also set to sheet style
    /// underline.
    /// </remarks>
    /// <param name="address">Address of the cell.</param>
    /// <returns><c>true</c> if there was such hyperlink and was deleted.
    ///   <c>false</c> otherwise.</returns>
    bool Delete(IXLAddress address);

    /// <summary>
    /// Get a hyperlink for a single cell.
    /// </summary>
    /// <param name="address">Address of the cell.</param>
    /// <exception cref="KeyNotFoundException">Cell doesn't have a hyperlink.</exception>
    XLHyperlink Get(IXLAddress address);

    /// <summary>
    /// Get a hyperlink for a single cell.
    /// </summary>
    /// <param name="address">Address of the cell.</param>
    /// <param name="hyperlink">Found hyperlink.</param>
    /// <returns><c>true</c> if there was a hyperlink for <paramref name="address"/>, <c>false</c> otherwise.</returns>
    bool TryGet(IXLAddress address, out XLHyperlink hyperlink);
}
