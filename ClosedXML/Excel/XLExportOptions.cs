namespace ClosedXML.Excel;

/// <summary>
/// options for exporting data to an Excel file.
/// </summary>
public class XLExportOptions
{
    /// <summary>
    /// Gets or sets the options for defining the appearance and behavior of cells within a column when exporting data.
    /// </summary>
    public XLExportCellOptions? Column { get; set; }

    /// <summary>
    /// Gets or sets the pattern used to remove striped rows from the exported data.
    /// </summary>
    public bool RemoveRowStriped { get; set; }

    /// <summary>
    /// Gets or sets the options for defining the appearance and behavior of cells within a row when exporting data.
    /// </summary>
    public XLExportCellOptions? Row { get; set; }

    /// <summary>
    /// Gets or sets the name of the worksheet in the Excel file.
    /// </summary>
    public string? SheetName { get; set; }
}
