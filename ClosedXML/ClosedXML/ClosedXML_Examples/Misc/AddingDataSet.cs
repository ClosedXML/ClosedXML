using System;
using System.Data;
using ClosedXML.Excel;

namespace ClosedXML_Examples.Misc
{
    public class AddingDataSet
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();

            var dataSet = GetDataSet();

            // Add all DataTables in the DataSet as a worksheets
            wb.Worksheets.Add(dataSet);

            wb.SaveAs(filePath);
        }

        // Private
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

            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
        // Override


        #endregion
    }
}
