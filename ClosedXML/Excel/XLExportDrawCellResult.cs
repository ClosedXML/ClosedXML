namespace ClosedXML.Excel;

/// <summary>
/// Represents the result of drawing a cell during the process of exporting data to an Excel file.
/// </summary>
public class XLExportDrawCellResult
{
    /// <summary>
    /// Gets or sets a value indicating whether the cell should be skipped during the export process.
    /// </summary>
    public bool IsSkip { get; set; }

    /// <summary>
    /// Gets or sets the options for defining the appearance and behavior of the cell.
    /// </summary>
    public XLExportCellOptions? Options { get; set; }

    /// <summary>
    /// Gets or sets the value to be displayed in the cell.
    /// </summary>
    public object? Value { get; set; }
}
