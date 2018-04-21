using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumnsCollection : IDictionary<Int32, XLColumn>
    {
        public void ShiftColumnsRight(Int32 startingColumn, Int32 columnsToShift)
        {
            foreach (var co in _dictionary.Keys.Where(k => k >= startingColumn).OrderByDescending(k => k))
            {
                var columnToMove = _dictionary[co];
                _dictionary.Remove(co);
                Int32 newColumnNum = co + columnsToShift;
                if (newColumnNum <= XLHelper.MaxColumnNumber)
                {
                    columnToMove.SetColumnNumber(newColumnNum);
                    _dictionary.Add(newColumnNum, columnToMove);
                }
            }
        }

        private readonly Dictionary<Int32, XLColumn> _dictionary = new Dictionary<Int32, XLColumn>();

        public void Add(int key, XLColumn value)
        {
            _dictionary.Add(key, value);
        }

        public bool ContainsKey(int key)
        {
            return _dictionary.ContainsKey(key);
        }

        public ICollection<int> Keys
        {
            get { return _dictionary.Keys; }
        }

        public bool Remove(int key)
        {
            return _dictionary.Remove(key);
        }

        public bool TryGetValue(int key, out XLColumn value)
        {
            return _dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLColumn> Values
        {
            get { return _dictionary.Values; }
        }

        public XLColumn this[int key]
        {
            get
            {
                return _dictionary[key];
            }
            set
            {
                _dictionary[key] = value;
            }
        }

        public void Add(KeyValuePair<int, XLColumn> item)
        {
            _dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            _dictionary.Clear();
        }

        public bool Contains(KeyValuePair<int, XLColumn> item)
        {
            return _dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, XLColumn>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return _dictionary.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(KeyValuePair<int, XLColumn> item)
        {
            return _dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<int, XLColumn>> GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        public void RemoveAll(Func<XLColumn, Boolean> predicate)
        {
            _dictionary.RemoveAll(predicate);
        }
    }
}
