using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class DataTableReader : IInsertDataReader
    {
        private readonly IEnumerable<DataRow> _dataRows;
        private readonly DataTable _dataTable;

        public DataTableReader(DataTable dataTable)
        {
            _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));
            _dataRows = _dataTable.Rows.Cast<DataRow>();
        }

        public DataTableReader(IEnumerable<DataRow> dataRows)
        {
            _dataRows = dataRows ?? throw new ArgumentNullException(nameof(dataRows));
            _dataTable = _dataRows.FirstOrDefault()?.Table;
        }

        public IEnumerable<IEnumerable<object>> GetData()
        {
            return _dataRows.Select(r => r.ItemArray);
        }

        public int GetPropertiesCount()
        {
            if (_dataTable != null)
                return _dataTable.Columns.Count;

            if (_dataRows.Any())
                return _dataRows.First().ItemArray.Length;

            return 0;
        }

        public string GetPropertyName(int propertyIndex)
        {
            if (propertyIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(propertyIndex), "Property index must be non-negative");

            if (_dataTable == null)
                return null;

            if (propertyIndex >= _dataTable.Columns.Count)
                throw new ArgumentOutOfRangeException($"{propertyIndex} exceeds the number of the table columns");

            return _dataTable.Columns[propertyIndex].Caption;
        }

        public int GetRecordsCount()
        {
            return _dataRows.Count();
        }
    }
}
