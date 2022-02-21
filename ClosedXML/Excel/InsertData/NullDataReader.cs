// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class NullDataReader : IInsertDataReader
    {
        private readonly int _count;

        public NullDataReader(IEnumerable<object> nulls)
        {
            _count = nulls.Count();
        }

        public IEnumerable<IEnumerable<object>> GetData()
        {
            var res = new object[] { null }.AsEnumerable();
            for (int i = 0; i < _count; i++)
            {
                yield return res;
            }
        }

        public int GetPropertiesCount()
        {
            return 0;
        }

        public string GetPropertyName(int propertyIndex)
        {
            return null;
        }

        public int GetRecordsCount()
        {
            return _count;
        }
    }
}
