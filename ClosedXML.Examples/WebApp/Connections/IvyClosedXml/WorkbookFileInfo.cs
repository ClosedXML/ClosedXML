using ClosedXML.Excel;

/// <summary>
/// Represents a workbook file with an associated XLWorkbook instance.
/// This class implements IDisposable to ensure proper disposal of the workbook.
/// </summary>
public class WorkbookFileInfo : IDisposable
{
    /// <summary>
    /// Initializes a new instance of the <see cref="WorkbookFileInfo"/> class with the specified file name.
    /// Creates a new XLWorkbook and adds a default worksheet named "Sheet1".
    /// </summary>
    /// <param name="fileName">The name of the workbook file.</param>
    public WorkbookFileInfo(string fileName)
    {
        FileName = fileName;
        Workbook = new XLWorkbook();
        Workbook.AddWorksheet("Sheet1");
    }

    /// <summary>
    /// Gets the file name associated with the workbook.
    /// </summary>
    public string FileName { get; }

    /// <summary>
    /// Gets the XLWorkbook instance.
    /// </summary>
    public XLWorkbook Workbook { get; }

    /// <summary>
    /// Disposes of the workbook and suppresses finalization.
    /// </summary>
    public void Dispose()
    {
        Workbook?.Dispose();
        GC.SuppressFinalize(this);
    }
}