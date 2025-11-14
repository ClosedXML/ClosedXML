using System;

namespace ClosedXML.Excel;

public interface IXLProtection : IEquatable<IXLProtection>
{
    /// <summary>
    /// Get or set whether a cell is locked.
    /// </summary>
    /// <remarks>
    /// When a cell is locked and its sheet is protected (<c>IXLWorksheet.Protection</c>), then
    /// the protected operations (<see cref="XLSheetProtectionElements"/>) are prohibited on
    /// the cell.
    /// </remarks>
    Boolean Locked { get; set; }

    /// <summary>
    /// Get or set whether cell is a hidden.
    /// </summary>
    /// <value>
    /// When a cell is hidden and its sheet is protected (<c>IXLWorksheet.Protection</c>), then
    /// application will display content of a cell in a grid, but it will hide formulas bar
    /// content (i.e. formulas or even plain text).
    /// </value>
    Boolean Hidden { get; set; }

    /// <summary>
    /// Set cell as locked.
    /// </summary>
    /// <inheritdoc cref="Locked"/>
    IXLStyle SetLocked();

    /// <summary>
    /// Set whether a cell is locked.
    /// </summary>
    /// <inheritdoc cref="Locked"/>
    IXLStyle SetLocked(Boolean value);

    /// <summary>
    /// Set cell as hidden.
    /// </summary>
    /// <inheritdoc cref="Hidden"/>
    IXLStyle SetHidden();

    /// <summary>
    /// Set whether a cell is hidden.
    /// </summary>
    /// <inheritdoc cref="Hidden"/>
    IXLStyle SetHidden(Boolean value);
}
