#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class SimpleTypeReader : IInsertDataReader
    {
        private readonly IEnumerable<object> _data;
        private readonly Type _itemType;

        public SimpleTypeReader(IEnumerable data)
        {
            _data = data?.Cast<object>() ?? throw new ArgumentNullException(nameof(data));
            _itemType = data.GetItemType();
        }

        public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
        {
            return _data.Select(item => new[] { item }.Select(XLCellValue.FromInsertedObject));
        }

        public int GetPropertiesCount()
        {
            return 1;
        }

        public string GetPropertyName(int propertyIndex = 0)
        {
            if (propertyIndex != 0)
                throw new ArgumentException("SimpleTypeReader supports only a single property");

            return _itemType.Name;
        }
    }
}
