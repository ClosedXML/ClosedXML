namespace ClosedXML.Excel;

/// <summary>
/// Listener for components that need to be notified about structural changes of a workbook
/// (adding/removing sheet, renaming). See <see cref="ISheetListener"/> for similar listener about
/// structural changes of a sheet.
/// </summary>
internal interface IWorkbookListener
{
    /// <summary>
    /// Method is called when sheet has already been renamed.
    /// </summary>
    /// <param name="oldSheetName">Old sheet name.</param>
    /// <param name="newSheetName">New sheet name, different from old one.</param>
    void OnSheetRenamed(string oldSheetName, string newSheetName);
}
