using System.Data;

namespace ClosedXmlExample.Apps;

/// <summary>
/// Workbook Manager App - Modern blade-based interface for managing Excel workbook files
/// </summary>
[App(icon: Icons.Table, title: "Workbooks Editor")]
public class ExcelEditorApp : ViewBase
{
    public override object? Build()
    {
        return this.UseBlades(() => new WorkbooksListBlade(), "Workbooks Editor", Size.Units(75));
    }
}

public record WorkbookListRecord(string FileName, int ColumnCount);

public class WorkbooksListBlade : ViewBase
{
    public override object? Build()
    {
        var blades = this.UseContext<IBladeController>();
        var workbookRepository = this.UseService<WorkbookRepository>();
        var refreshToken = this.UseRefreshToken();

        this.UseEffect(() =>
        {
            if (refreshToken.ReturnValue is string fileName)
            {
                blades.Pop(this, true);
                blades.Push(this, new WorkbookEditorBlade(fileName), fileName, width: Size.Units(100));
            }
        }, [refreshToken]);

        var onItemClicked = new Action<Event<ListItem>>(e =>
        {
            var workbook = (WorkbookListRecord)e.Sender.Tag!;
            blades.Push(this, new WorkbookEditorBlade(workbook.FileName), workbook.FileName, width: Size.Units(100));
        });

        ListItem CreateItem(WorkbookListRecord record) =>
            new(
                title: record.FileName,
                onClick: onItemClicked,
                tag: record
            );

        var createBtn = Icons.Plus.ToButton(_ =>
        {
            blades.Pop(this);
        }).ToTrigger((isOpen) => new WorkbookCreateDialog(isOpen, refreshToken));

        return new FilteredListView<WorkbookListRecord>(
            fetchRecords: (filter) => FetchWorkbooks(workbookRepository, filter),
            createItem: CreateItem,
            toolButtons: createBtn,
            onFilterChanged: _ =>
            {
                blades.Pop(this);
            }
        );
    }

    private async Task<WorkbookListRecord[]> FetchWorkbooks(WorkbookRepository repository, string filter)
    {
        return await Task.Run(() =>
        {
            var files = repository.GetFiles();

            var filtered = string.IsNullOrWhiteSpace(filter)
                ? files
                : files.Where(f => f.FileName.Contains(filter, StringComparison.OrdinalIgnoreCase));

            return filtered
                .Select(f => new WorkbookListRecord(
                    f.FileName, 
                    f.Workbook.Worksheets.FirstOrDefault()?.ColumnsUsed().Count() ?? 0))
                .ToArray();
        });
    }
}

/// <summary>
/// Workbook Editor Blade - Shows tabs for each worksheet with full editing capabilities
/// </summary>
public class WorkbookEditorBlade(string fileName) : ViewBase
{
    public override object? Build()
    {
        var blades = this.UseContext<IBladeController>();
        var workbookRepository = this.UseService<WorkbookRepository>();
        var refreshToken = this.UseRefreshToken();

        // Set current file
        this.UseEffect(() =>
        {
            workbookRepository.SetCurrentFile(fileName);
        }, [EffectTrigger.AfterInit()]);
        // Get the first worksheet - passing fileName directly to avoid shared state issues
        var table = workbookRepository.GetCurrentTable(fileName);

        return new WorksheetEditor(table, fileName, blades);
    }
}

/// <summary>
/// Worksheet Editor - Allows adding columns, rows, and saving changes
/// </summary>
public class WorksheetEditor(DataTable table, string fileName, IBladeController blades) : ViewBase
{
    public override object? Build()
    {
        var client = this.UseService<IClientProvider>();
        var workbookRepository = this.UseService<WorkbookRepository>();
        var refreshToken = this.UseRefreshToken();

        var columnName = this.UseState<string?>(() => null);
        var columnTypes = new string[] { "int", "double", "decimal", "long", "string" };
        var selectedType = this.UseState("string");

        var addColumnButton = new Button("Add Column", _ =>
        {
            if (string.IsNullOrWhiteSpace(columnName.Value))
            {
                client.Toast("Column name cannot be empty");
                return;
            }

            try
            {
                var colName = columnName.Value;
                table.Columns.Add(colName, GetColumnTypeFromString(selectedType.Value));
                
                // Auto-save changes - passing fileName directly to avoid shared state issues
                workbookRepository.Save(fileName, table);
                
                columnName.Set(String.Empty);
                refreshToken.Refresh();
                client.Toast($"Column '{colName}' added and saved!");
            }
            catch (Exception ex)
            {
                client.Toast($"Error: {ex.Message}");
            }
        })
        .Variant(ButtonVariant.Primary)
        .Icon(Icons.Plus);

        var columnEditorCard = new Card(
            Layout.Vertical().Gap(2)
            | Layout.Horizontal().Gap(2)
                | columnName.ToTextInput(placeholder: "Column name")
                | selectedType.ToSelectInput(columnTypes.ToOptions()).Variant(SelectInputs.Select)
                | addColumnButton
        ).Title($"Add New Column");

        var rowEditor = new RowEditor(table, refreshToken, workbookRepository, fileName);
        var tableView = new WorksheetTableView(table);

        var deleteBtn = new Button("Delete", onClick: e =>
            {
                workbookRepository.RemoveFile(fileName);
                blades.Pop(refresh: true);
            })
            .Variant(ButtonVariant.Destructive)
            .Icon(Icons.Trash)
            .WithConfirm($"Are you sure you want to delete workbook '{fileName}'?", "Delete Workbook");

        var columnCount = table.Columns.Count;
        var rowCount = table.Rows.Count;

        var footer = new Card(
            Layout.Vertical().Gap(2)
            | Layout.Horizontal().Gap(4)
                | Text.Label($"{columnCount} columns")
                | Text.Label($"{rowCount} rows")
            | deleteBtn
        ).Title(fileName);

        return Layout.Vertical().Gap(4)
            | columnEditorCard
            | new Card(rowEditor).Title("Add New Row")
            | new Card(tableView).Title("Table Data")
            | footer;
    }

    private static Type GetColumnTypeFromString(string selectedType)
    {
        return selectedType switch
        {
            "string" => typeof(string),
            "int" => typeof(int),
            "double" => typeof(double),
            "decimal" => typeof(decimal),
            "long" => typeof(long),
            _ => throw new ArgumentOutOfRangeException(selectedType, "Column type not supported!")
        };
    }
}

/// <summary>
/// Row Editor - UI for adding new rows to the table
/// </summary>
public class RowEditor(DataTable table, RefreshToken refreshToken, WorkbookRepository workbookRepository, string fileName) : ViewBase
{
    public override object? Build()
    {
        var client = this.UseService<IClientProvider>();
        var inputsForRowData = new List<IState<string?>>();
        var dataColumns = table.Columns.Cast<DataColumn>().ToList();

        if (dataColumns.Count == 0)
        {
            return Text.Small("Add columns first to start adding rows");
        }

        var inputFields = new List<object>();
        foreach (var col in dataColumns)
        {
            var inputState = this.UseState<string?>(() => null);
            inputsForRowData.Add(inputState);
            inputFields.Add(inputState.ToTextInput(placeholder: col.ColumnName));
        }

        var addRowButton = new Button("Add Row", _ =>
        {
            try
            {
                var newRow = inputsForRowData.Select(input => input.Value ?? "").ToArray();
                table.Rows.Add(newRow);
                
                // Auto-save changes - passing fileName directly to avoid shared state issues
                workbookRepository.Save(fileName, table);
                
                // Clear inputs
                foreach (var input in inputsForRowData)
                {
                    input.Set(String.Empty);
                }
                
                refreshToken.Refresh();
                client.Toast("Row added and saved!");
            }
            catch (Exception ex)
            {
                client.Toast($"Error: {ex.Message}");
            }
        })
        .Variant(ButtonVariant.Primary)
        .Icon(Icons.Plus);

        return Layout.Vertical().Gap(2)
            | Layout.Horizontal().Gap(2) | inputFields.ToArray()
            | addRowButton;
    }
}

/// <summary>
/// Worksheet Table View - Displays the table data
/// </summary>
public class WorksheetTableView(DataTable table) : ViewBase
{
    public override object? Build()
    {
        if (table.Columns.Count == 0)
        {
            return Text.Small("No columns defined yet");
        }

        var rows = table.AsEnumerable().ToList();
        var tableColumns = table.Columns.Cast<DataColumn>().ToList();
        var listOfTableRows = new List<TableRow>();

        // Header row
        var headerCells = tableColumns.Select(col => new TableCell(col.ColumnName).IsHeader()).ToList();
        listOfTableRows.Add(new TableRow([.. headerCells]));

        // Data rows
        foreach (var row in rows)
        {
            var dataCells = row.ItemArray.Select(value => new TableCell(value?.ToString() ?? "")).ToList();
            listOfTableRows.Add(new TableRow([.. dataCells]));
        }

        // Always return table, even if there are only headers (no data rows)
        return new Table([.. listOfTableRows]);
    }
}

public record WorkbookCreateRequest
{
    [Required]
    public string FileName { get; init; } = "";
}

public class WorkbookCreateDialog(IState<bool> isOpen, RefreshToken refreshToken) : ViewBase
{
    public override object? Build()
    {
        var workbookRepository = this.UseService<WorkbookRepository>();
        var client = this.UseService<IClientProvider>();
        var workbook = this.UseState(() => new WorkbookCreateRequest());

        this.UseEffect(() =>
        {
            try
            {
                CreateWorkbook(workbookRepository, workbook.Value);
                refreshToken.Refresh(workbook.Value.FileName);
            }
            catch (Exception ex)
            {
                client.Toast($"Error: {ex.Message}");
                throw;
            }
        }, [workbook]);

        return workbook
            .ToForm()
            .ToDialog(isOpen, title: "Create Workbook", submitTitle: "Create");
    }

    private void CreateWorkbook(WorkbookRepository repository, WorkbookCreateRequest request)
    {
        repository.AddNewFile(request.FileName);
    }
}

