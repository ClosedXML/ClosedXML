using ClosedXML.Excel;
using System;
using System.Data;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class AddingDataSet : IXLExample
    {
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();

            var dataSet = GetDataSet();

            // Add all DataTables in the DataSet as a worksheets
            wb.Worksheets.Add(dataSet);

            foreach (var ws in wb.Worksheets)
                ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        private DataSet GetDataSet()
        {
            var ds = new DataSet();
            ds.Tables.Add(GetTable("Patients"));
            ds.Tables.Add(GetTable("Employees"));
            ds.Tables.Add(GetTable("Information"));
            return ds;
        }

        private DataTable GetTable(String tableName)
        {
            DataTable table = new DataTable();
            table.TableName = tableName;
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            return table;
        }
    }
}
