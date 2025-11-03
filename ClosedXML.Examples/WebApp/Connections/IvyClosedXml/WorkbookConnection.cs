using Ivy.Connections;
using System.Data;

/// <summary>
/// Implement IConnection interface so that WorkbookConnection is registered in the DI container
/// </summary>
public class WorkbookConnection : IConnection
{
    // Static repository instance - single shared database for the entire application
    private static readonly WorkbookRepository _sharedRepository = new WorkbookRepository();
    
    public WorkbookConnection()
    {
        InitializeSampleData();
    }
    
    /// <summary>
    /// Creates sample workbooks with example data for demonstration
    /// </summary>
    private void InitializeSampleData()
    {
        // Check if already initialized to avoid duplicating data on each connection creation
        if (_sharedRepository.GetFiles().Count > 0)
            return;
            
        try
        {
            // Sample 1: Employees
            _sharedRepository.AddNewFile("Employees.xlsx");
            var employeesTable = new DataTable { TableName = "Employees" };
            employeesTable.Columns.Add("Name", typeof(string));
            employeesTable.Columns.Add("Position", typeof(string));
            employeesTable.Columns.Add("Salary", typeof(decimal));
            employeesTable.Columns.Add("Department", typeof(string));
            
            employeesTable.Rows.Add("John Doe", "Developer", 75000, "IT");
            employeesTable.Rows.Add("Jane Smith", "Manager", 85000, "HR");
            employeesTable.Rows.Add("Mike Johnson", "Designer", 65000, "Marketing");
            employeesTable.Rows.Add("Sarah Williams", "Analyst", 70000, "Finance");
            
            _sharedRepository.Save("Employees.xlsx", employeesTable);
            
            // Sample 2: Products
            _sharedRepository.AddNewFile("Products.xlsx");
            var productsTable = new DataTable { TableName = "Products" };
            productsTable.Columns.Add("Product", typeof(string));
            productsTable.Columns.Add("Price", typeof(decimal));
            productsTable.Columns.Add("Stock", typeof(int));
            productsTable.Columns.Add("Category", typeof(string));
            
            productsTable.Rows.Add("Laptop", 1299.99, 45, "Electronics");
            productsTable.Rows.Add("Mouse", 29.99, 150, "Electronics");
            productsTable.Rows.Add("Desk", 399.99, 20, "Furniture");
            productsTable.Rows.Add("Chair", 249.99, 35, "Furniture");
            productsTable.Rows.Add("Monitor", 349.99, 60, "Electronics");
            
            _sharedRepository.Save("Products.xlsx", productsTable);
            
            // Sample 3: Sales
            _sharedRepository.AddNewFile("Sales.xlsx");
            var salesTable = new DataTable { TableName = "Sales" };
            salesTable.Columns.Add("Date", typeof(string));
            salesTable.Columns.Add("Product", typeof(string));
            salesTable.Columns.Add("Quantity", typeof(int));
            salesTable.Columns.Add("Total", typeof(decimal));
            
            salesTable.Rows.Add("2024-01-15", "Laptop", 3, 3899.97);
            salesTable.Rows.Add("2024-01-16", "Mouse", 12, 359.88);
            salesTable.Rows.Add("2024-01-17", "Chair", 5, 1249.95);
            
            _sharedRepository.Save("Sales.xlsx", salesTable);
        }
        catch (Exception)
        {
            // If samples already exist or error occurs, just continue
        }
    }

    public WorkbookRepository GetWorkbookRepository()
    {
        return _sharedRepository;
    }

    public string GetConnectionType()
    {
        return typeof(WorkbookConnection).ToString();
    }

    public string GetContext(string connectionPath)
    {
        throw new NotImplementedException();
    }

    public ConnectionEntity[] GetEntities()
    {
        throw new NotImplementedException();
    }

    public string GetName() => nameof(WorkbookConnection);

    public string GetNamespace() => typeof(WorkbookConnection).Namespace;

    public void RegisterServices(IServiceCollection services)
    {
        // Register the static WorkbookRepository instance as Singleton - single shared database for the entire application
        services.AddSingleton<WorkbookRepository>(_ => _sharedRepository);
        services.AddSingleton<WorkbookConnection>();
    }
}