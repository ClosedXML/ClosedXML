using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLIdManager
    {
        private HashSet<Int32> _hash = new HashSet<Int32>();
        

        public Int32 GetNext()
        {
            if (_hash.Count == 0)
            {
                _hash.Add(1);
                return 1;
            }

            Int32 id = 1;
            while (true)
            {
                if (!_hash.Contains(id))
                {
                    _hash.Add(id);
                    return id;
                }
                id++;
            }
        }
        public void Add(Int32 value)
        {
            _hash.Add(value);
        }
        public void Add(IEnumerable<Int32> values)
        {
            values.ForEach(v => _hash.Add(v));
        }
    }
}
