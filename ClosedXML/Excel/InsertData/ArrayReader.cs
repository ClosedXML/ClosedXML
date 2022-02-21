// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class ArrayReader : IInsertDataReader
    {
        private readonly IEnumerable<IEnumerable> _data;

        public ArrayReader(IEnumerable<IEnumerable> data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        public IEnumerable<IEnumerable<object>> GetData()
        {
            return _data.Select(item => item.Cast<object>());
        }

        public int GetPropertiesCount()
        {
            if (!_data.Any())
                return 0;

            return _data.First().Cast<object>().Count();
        }

        public string GetPropertyName(int propertyIndex)
        {
            return null;
        }

        public int GetRecordsCount()
        {
            return _data.Count();
        }
    }
}
