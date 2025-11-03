# ClosedXML

<img width="1911" height="909" alt="image" src="https://github.com/user-attachments/assets/1404dda9-a66c-4f80-8f40-85fa0e9a2bbe" />

<img width="1916" height="914" alt="image" src="https://github.com/user-attachments/assets/cebe4ce4-b28a-4c1a-8f7d-5eeccc89a6b8" />

## One-Click Development Environment

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://github.com/codespaces/new?hide_repo_select=true&ref=main&repo=Ivy-Interactive%2FIvy-Examples&machine=standardLinux32gb&devcontainer_path=.devcontainer%2Fclosedxml%2Fdevcontainer.json&location=EuropeWest)

Click the badge above to open Ivy Examples repository in GitHub Codespaces with:
- **.NET 9.0** SDK pre-installed
- **Ready-to-run** development environment
- **No local setup** required

## Created Using Ivy

Web application created using [Ivy-Framework](https://github.com/Ivy-Interactive/Ivy-Framework).

**Ivy** - The ultimate framework for building internal tools with LLM code generation by unifying front-end and back-end into a single C# codebase. With Ivy, you can build robust internal tools and dashboards using C# and AI assistance based on your existing database.

Ivy is a web framework for building interactive web applications using C# and .NET.

## Interactive Example For Excel File Management

This example demonstrates Excel workbook management using the [ClosedXML library](https://github.com/ClosedXML/ClosedXML) integrated with Ivy. ClosedXML is a .NET library for reading, manipulating, and writing Excel 2007+ (.xlsx, .xlsm) files without requiring Excel to be installed.

**What This Application Does:**

This implementation creates two integrated applications for comprehensive Excel file management:

### Workbooks Viewer
- **File Selection**: Choose from available workbooks via dropdown menu
- **Data Preview**: View table data with automatic row limiting for performance
- **Refresh Button**: Update file list to see newly created workbooks
- **Clean Interface**: Split-panel layout with file selection and data preview

### Workbooks Editor
- **Blade-Based Interface**: Modern multi-panel interface for managing multiple workbooks
- **Create Workbooks**: Add new Excel files with custom names
- **Add Columns**: Create new columns with type selection (string, int, double, decimal, long)
- **Add Rows**: Insert data rows with automatic validation
- **Auto-Save**: Changes are automatically saved to in-memory workbooks
- **Delete Workbooks**: Remove files with confirmation dialog
- **Real-Time Updates**: See changes reflected immediately in the interface

**Technical Implementation:**

- Uses ClosedXML's `XLWorkbook` class for Excel file manipulation
- Implements in-memory workbook storage via singleton `WorkbookRepository`
- Converts Excel tables to `DataTable` for easy manipulation
- Handles multiple file formats and column types
- Creates blade-based navigation for seamless multi-file editing
- Implements sample data initialization with Employees, Products, and Sales workbooks
- Uses Ivy's reactive state management for real-time UI updates
- Integrates refresh tokens for manual data synchronization
- Built with two main C# views:
  - `Apps/WorkbooksViewerApp.cs` - Simple viewer for data preview
  - `Apps/WorkbooksEditorApp.cs` - Full-featured editor with blade navigation

**Key Features:**

- **Type-Safe Column Creation**: Support for multiple data types with automatic conversion
- **Empty Row Filtering**: Automatically removes empty rows from imported data
- **Singleton Pattern**: Shared repository ensures consistent data across all app instances
- **Validation**: File name validation with invalid character detection
- **Worksheet Management**: Automatic table creation and management within worksheets
- **In-Memory Storage**: No physical files created - all data stored in memory for fast access

## How to Run

1. **Prerequisites**: .NET 9.0 SDK
2. **Navigate to the example**:
   ```bash
   cd closedxml
   ```
3. **Restore dependencies**:
   ```bash
   dotnet restore
   ```
4. **Run the application**:
   ```bash
   dotnet watch
   ```
5. **Open your browser** to the URL shown in the terminal (typically `http://localhost:5010`)

## How to Deploy

Deploy this example to Ivy's hosting platform:

1. **Navigate to the example**:
   ```bash
   cd closedxml
   ```
2. **Deploy to Ivy hosting**:
   ```bash
   ivy deploy
   ```
This will deploy your Excel workbook management application with a single command.

## Learn More

- ClosedXML GitHub repository: [github.com/ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)
- Ivy Documentation: [docs.ivy.app](https://docs.ivy.app)

