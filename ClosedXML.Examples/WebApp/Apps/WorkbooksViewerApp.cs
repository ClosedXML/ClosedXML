using System.Data;

namespace ClosedXmlExample.Apps;

/// <summary>
/// Workbooks Viewer App - Default view with file selection and data preview
/// </summary>
[App(icon: Icons.FileSpreadsheet, title: "Workbooks Viewer")]
public class WorkbooksViewerApp : ViewBase
{
    public override object? Build()
    {
        var workbookRepository = this.UseService<WorkbookRepository>();
        var refreshToken = this.UseRefreshToken();
        
        var files = workbookRepository.GetFiles();
        var selectedFileIndex = this.UseState(0);
        
        // Get selected file data
        DataTable? selectedTable = null;
        WorkbookFileInfo? selectedFile = null;
        
        if (files.Count > 0 && selectedFileIndex.Value < files.Count)
        {
            selectedFile = files[selectedFileIndex.Value];
            try
            {
                selectedTable = workbookRepository.GetCurrentTable(selectedFile.FileName);
            }
            catch
            {
                // Handle error silently
            }
        }
        
        // Dropdown menu items for file selection
        var fileMenuItems = files
            .Select((file, idx) => MenuItem.Default(file.FileName).HandleSelect(() => selectedFileIndex.Value = idx))
            .ToArray();
        
        var fileDropDown = new Button(selectedFile?.FileName ?? "Select File")
            .Primary()
            .Icon(Icons.ChevronDown)
            .WithDropDown(fileMenuItems);
        
        var refreshButton = Icons.RefreshCw.ToButton(_ => refreshToken.Refresh())
            .Variant(ButtonVariant.Secondary)
            .WithTooltip("Refresh file list");
        
        // Left Card - File Selection and Info
        var columnCount = selectedTable?.Columns.Count ?? 0;
        var rowCount = selectedTable?.Rows.Count ?? 0;
        
        var leftCard = new Card(
            Layout.Vertical().Gap(4).Padding(2)
            | Text.H2("File Selection")
            | Text.Muted("Choose a workbook to preview")
            | (Layout.Horizontal().Gap(2)
                | fileDropDown
                | refreshButton)
            | new Spacer()
            | Text.Small("This demo uses the ClosedXML NuGet package to work with Excel files.")
            | Text.Markdown("Built with [Ivy Framework](https://github.com/Ivy-Interactive/Ivy-Framework) and [ClosedXML](https://github.com/ClosedXML/ClosedXML)")
        ).Width(Size.Fraction(0.45f)).Height(110);
        
        // Right Card - Data Table
        var tableContent = selectedTable != null && selectedTable.Columns.Count > 0
            ? BuildDataTable(selectedTable)
            : Text.Muted("No data to display");
        
        var rightCard = new Card(
            Layout.Vertical().Gap(4).Padding(2)
            | Text.H2("Data Preview")
            | Text.Muted(selectedFile != null ? $"Viewing: {selectedFile.FileName}" : "Select a file")
            | tableContent
        ).Width(Size.Fraction(0.45f)).Height(Size.Fit().Min(110));
        
        return Layout.Horizontal().Gap(6).Align(Align.Center)
            | leftCard
            | rightCard;
    }
    
    private object BuildDataTable(DataTable table)
    {
        var rows = table.AsEnumerable().ToList();
        var tableColumns = table.Columns.Cast<DataColumn>().ToList();
        var listOfTableRows = new List<TableRow>();
        
        // Header row
        var headerCells = tableColumns.Select(col => 
            new TableCell(col.ColumnName).IsHeader()
        ).ToList();
        listOfTableRows.Add(new TableRow([.. headerCells]));
        
        // Data rows - limit to first 50 rows for performance
        var displayRows = rows.Take(50).ToList();
        foreach (var row in displayRows)
        {
            var dataCells = row.ItemArray.Select(value => 
                new TableCell(value?.ToString() ?? "")
            ).ToList();
            listOfTableRows.Add(new TableRow([.. dataCells]));
        }
        
        var tableView = new Table([.. listOfTableRows]);
        
        // Show message if there are more rows
        if (rows.Count > 50)
        {
            return Layout.Vertical().Gap(2)
                | tableView
                | Text.Small($"Showing first 50 of {rows.Count} rows");
        }
        
        return tableView;
    }
}

