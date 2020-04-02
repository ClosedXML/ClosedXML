// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class UntypedObjectReader : IInsertDataReader
    {
        private readonly IEnumerable<object> _data;
        private readonly IEnumerable<IInsertDataReader> _readers;

        public UntypedObjectReader(IEnumerable data)
        {
            _data = (data ?? new object[0]).Cast<object>();
            _readers = CreateReaders().ToList();

            IEnumerable<IInsertDataReader> CreateReaders()
            {
                if (!_data.Any())
                    yield break;

                List<object> itemsOfSameType = new List<object>();
                Type previousType = null;

                foreach (var item in _data)
                {
                    var currentType = item?.GetType();

                    if (previousType != currentType && itemsOfSameType.Count > 0)
                    {
                        yield return CreateReader(itemsOfSameType, previousType);
                        itemsOfSameType.Clear();
                    }
                    itemsOfSameType.Add(item);
                    previousType = currentType;
                }

                if (itemsOfSameType.Count > 0)
                {
                    yield return CreateReader(itemsOfSameType, previousType);
                }
            }

            IInsertDataReader CreateReader(List<object> itemsOfSameType, Type itemType)
            {
                if (itemType == null)
                    return new NullDataReader(itemsOfSameType);

                var items = Array.CreateInstance(itemType, itemsOfSameType.Count);
                Array.Copy(itemsOfSameType.ToArray(), items, items.Length);

                return InsertDataReaderFactory.Instance.CreateReader(items);
            }
        }

        public IEnumerable<IEnumerable<object>> GetData()
        {
            foreach (var reader in _readers)
            {
                foreach (var item in reader.GetData())
                {
                    yield return item;
                }
            }
        }

        public int GetPropertiesCount()
        {
            return GetFirstNonNullReader()?.GetPropertiesCount() ?? 0;
        }

        public string GetPropertyName(int propertyIndex)
        {
            return GetFirstNonNullReader()?.GetPropertyName(propertyIndex);
        }

        public int GetRecordsCount()
        {
            return _data.Count();
        }

        private IInsertDataReader GetFirstNonNullReader()
        {
            return _readers.FirstOrDefault(r => !(r is NullDataReader));
        }
    }
}
