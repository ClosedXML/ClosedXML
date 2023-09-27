#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class DataRecordReader : IInsertDataReader
    {
        private readonly IEnumerable<XLCellValue>[] _inMemoryData;
        private string[] _columns;

        public DataRecordReader(IEnumerable<IDataRecord> data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            _inMemoryData = ReadToEnd(data).ToArray();
        }

        public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
        {
            return _inMemoryData;
        }

        public int GetPropertiesCount()
        {
            return _columns.Length;
        }

        public string GetPropertyName(int propertyIndex)
        {
            if (propertyIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(propertyIndex), "Property index must be non-negative");

            if (_columns == null)
                return null;

            if (propertyIndex >= _columns.Length)
                throw new ArgumentOutOfRangeException($"{propertyIndex} exceeds the number of the table columns");

            return _columns[propertyIndex];
        }

        private IEnumerable<IEnumerable<XLCellValue>> ReadToEnd(IEnumerable<IDataRecord> data)
        {
            foreach (var dataRecord in data)
            {
                yield return ToEnumerable(dataRecord).ToArray();
            }
        }

        private IEnumerable<XLCellValue> ToEnumerable(IDataRecord dataRecord)
        {
            var firstRow = false;
            if (_columns == null)
            {
                firstRow = true;
                _columns = new string[dataRecord.FieldCount];
            }

            for (int i = 0; i < dataRecord.FieldCount; i++)
            {
                if (firstRow)
                    _columns[i] = dataRecord.GetName(i);

                var value = dataRecord[i];
                yield return XLCellValue.FromInsertedObject(value);
            }
        }
    }
}
