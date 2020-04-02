// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class InsertDataReaderFactory
    {
        private static readonly Lazy<InsertDataReaderFactory> _instance =
            new Lazy<InsertDataReaderFactory>(() => new InsertDataReaderFactory());

        public static InsertDataReaderFactory Instance => _instance.Value;

        public IInsertDataReader CreateReader(IEnumerable data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            var itemType = data.GetItemType();

            if (itemType == null || itemType == typeof(Object))
                return new UntypedObjectReader(data);
            else if (itemType.IsNullableType() && itemType.GetUnderlyingType().IsSimpleType())
                return new SimpleNullableTypeReader(data);
            else if (itemType.IsSimpleType())
                return new SimpleTypeReader(data);
            else if (typeof(IDataRecord).IsAssignableFrom(itemType))
                return new DataRecordReader(data.OfType<IDataRecord>());
            else if (itemType.IsArray || typeof(IEnumerable).IsAssignableFrom(itemType))
                return new ArrayReader(data.Cast<IEnumerable>());
            else if (itemType == typeof(DataRow))
                return new DataTableReader(data.Cast<DataRow>());

            return new ObjectReader(data);
        }

        public IInsertDataReader CreateReader<T>(IEnumerable<T[]> data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            return new ArrayReader(data);
        }

        public IInsertDataReader CreateReader(IEnumerable<IEnumerable> data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            if (data?.GetType().GetElementType() == typeof(String))
                return new SimpleTypeReader(data);

            return new ArrayReader(data);
        }

        public IInsertDataReader CreateReader(DataTable dataTable)
        {
            if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));

            return new DataTableReader(dataTable);
        }
    }
}
