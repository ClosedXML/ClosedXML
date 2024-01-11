// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.InsertData
{
    internal class NullDataReader : IInsertDataReader
    {
        private readonly XLCellValue[] _row = { Blank.Value };
        private readonly int _count;

        public NullDataReader(IEnumerable<object> nulls)
        {
            _count = nulls.Count();
        }

        public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
        {
            return Enumerable.Repeat(_row, _count);
        }

        public int GetPropertiesCount()
        {
            return 0;
        }

        public string? GetPropertyName(int propertyIndex)
        {
            return null;
        }
    }
}
