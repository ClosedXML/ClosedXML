using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLIdManager
    {
        private readonly HashSet<int> _hash = new HashSet<int>();
        

        public int GetNext()
        {
            if (_hash.Count == 0)
            {
                _hash.Add(1);
                return 1;
            }

            var id = 1;
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
        public void Add(int value)
        {
            _hash.Add(value);
        }
        public void Add(IEnumerable<int> values)
        {
            values.ForEach(v => _hash.Add(v));
        }
    }
}
