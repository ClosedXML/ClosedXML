using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumnsCollection : IDictionary<Int32, XLColumn>, IDisposable
    {
        public void ShiftColumnsRight(Int32 startingColumn, Int32 columnsToShift)
        {
            foreach (var ro in _dictionary.Keys.Where(k => k >= startingColumn).OrderByDescending(k => k))
            {
                var columnToMove = _dictionary[ro];
                Int32 newColumnNum = ro + columnsToShift;
                if (newColumnNum <= XLHelper.MaxColumnNumber)
                {
                    var newColumn = new XLColumn(columnToMove)
                                        {
                                            RangeAddress =
                                            {
                                                FirstAddress = new XLAddress(1, newColumnNum, false, false),
                                                LastAddress =
                                                    new XLAddress(XLHelper.MaxRowNumber, newColumnNum, false, false)
                                            }
                                        };
                                        
                    _dictionary.Add(newColumnNum, newColumn);
                }
                _dictionary.Remove(ro);
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

        public void Dispose()
        {
            _dictionary.Values.ForEach(c=>c.Dispose());
        }
    }
}
