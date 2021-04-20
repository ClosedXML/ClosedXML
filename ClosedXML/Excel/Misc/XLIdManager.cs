using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLIdManager
    {
        private HashSet<Int32> _hash = new HashSet<Int32>();

        private int _nextId = 1;

        public Int32 GetNext()
        {
            Int32 id = _nextId;
            while (true)
            {
                if (_hash.Add(id))
                {
                    _nextId = id + 1;
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
            values.ForEach(v => { _hash.Add(v); });
        }
    }
}
