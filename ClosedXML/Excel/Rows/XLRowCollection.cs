using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRowsCollection : IDictionary<Int32, XLRow>
    {
        private readonly Dictionary<Int32, XLRow> _dictionary = new Dictionary<Int32, XLRow>();

        public Dictionary<Int32, XLRow> Deleted { get; } = new Dictionary<Int32, XLRow>();

        public Int32 MaxRowUsed;

        #region IDictionary<int,XLRow> Members

        public void Add(int key, XLRow value)
        {
            if (key > MaxRowUsed) MaxRowUsed = key;

            if (Deleted.ContainsKey(key))
                Deleted.Remove(key);

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
            if (!Deleted.ContainsKey(key))
                Deleted.Add(key, _dictionary[key]);

            return _dictionary.Remove(key);
        }

        public bool TryGetValue(int key, out XLRow value)
        {
            return _dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLRow> Values
        {
            get { return _dictionary.Values; }
        }

        public XLRow this[int key]
        {
            get { return _dictionary[key]; }
            set { _dictionary[key] = value; }
        }

        public void Add(KeyValuePair<int, XLRow> item)
        {
            if (item.Key > MaxRowUsed) MaxRowUsed = item.Key;

            if (Deleted.ContainsKey(item.Key))
                Deleted.Remove(item.Key);

            _dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            _dictionary.Clear();
        }

        public bool Contains(KeyValuePair<int, XLRow> item)
        {
            return _dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, XLRow>[] array, int arrayIndex)
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

        public bool Remove(KeyValuePair<int, XLRow> item)
        {
            if (!Deleted.ContainsKey(item.Key))
                Deleted.Add(item.Key, _dictionary[item.Key]);

            return _dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<int, XLRow>> GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        #endregion IDictionary<int,XLRow> Members

        public void ShiftRowsDown(Int32 startingRow, Int32 rowsToShift)
        {
            foreach (int ro in _dictionary.Keys.Where(k => k >= startingRow).OrderByDescending(k => k))
            {
                var rowToMove = _dictionary[ro];
                _dictionary.Remove(ro);
                Int32 newRowNum = ro + rowsToShift;
                if (newRowNum <= XLHelper.MaxRowNumber)
                {
                    rowToMove.SetRowNumber(newRowNum);
                    _dictionary.Add(newRowNum, rowToMove);
                }
            }
        }

        public void RemoveAll(Func<XLRow, Boolean> predicate)
        {
            foreach (var row in _dictionary.Values.Where(predicate).Where(row1 => !Deleted.ContainsKey(row1.RowNumber())))
            {
                Deleted.Add(row.RowNumber(), row);
            }

            _dictionary.RemoveAll(predicate);
        }
    }
}
