using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRowsCollection: IDictionary<Int32, XLRow>
    {
        public void ShiftRowsDown(Int32 startingRow, Int32 rowsToShift)
        {
            foreach (var ro in dictionary.Keys.Where(k => k >= startingRow).OrderByDescending(k => k))
            {
                var rowToMove = dictionary[ro];
                Int32 newRow = ro + rowsToShift;
                if (newRow <= XLWorksheet.MaxNumberOfRows)
                {
                    dictionary.Add(newRow, new XLRow(rowToMove));
                }
                dictionary.Remove(ro);
            }
        }

        private Dictionary<Int32, XLRow> dictionary = new Dictionary<Int32, XLRow>();
        private Dictionary<Int32, XLRow> deleted = new Dictionary<Int32, XLRow>();
        public Dictionary<Int32, XLRow> Deleted
        {
            get
            {
                return deleted;
            }
        }

        public void Add(int key, XLRow value)
        {
            if (deleted.ContainsKey(key))
                deleted.Remove(key);

            dictionary.Add(key, value);
        }

        public bool ContainsKey(int key)
        {
            return dictionary.ContainsKey(key);
        }

        public ICollection<int> Keys
        {
            get { return dictionary.Keys; }
        }

        public bool Remove(int key)
        {
            if (!deleted.ContainsKey(key))
                deleted.Add(key, dictionary[key]);

            return dictionary.Remove(key);
        }

        public bool TryGetValue(int key, out XLRow value)
        {
            return dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLRow> Values
        {
            get { return dictionary.Values; }
        }

        public XLRow this[int key]
        {
            get
            {
                return dictionary[key];
            }
            set
            {
                dictionary[key] = value;
            }
        }

        public void Add(KeyValuePair<int, XLRow> item)
        {
            if (deleted.ContainsKey(item.Key))
                deleted.Remove(item.Key);

            dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            foreach (var kp in dictionary)
            {
                if (!deleted.ContainsKey(kp.Key))
                    deleted.Add(kp.Key, kp.Value);
            }

            dictionary.Clear();
        }

        public bool Contains(KeyValuePair<int, XLRow> item)
        {
            return dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, XLRow>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return dictionary.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(KeyValuePair<int, XLRow> item)
        {
            if (!deleted.ContainsKey(item.Key))
                deleted.Add(item.Key, dictionary[item.Key]);

            return dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<int, XLRow>> GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        public void RemoveAll(Func<XLRow, Boolean> predicate)
        {
            foreach (var kp in dictionary.Values.Where(predicate).Select(c=>c))
            {
                if (!deleted.ContainsKey(kp.RowNumber()))
                    deleted.Add(kp.RowNumber(), kp);
            }

            dictionary.RemoveAll(predicate);
        }
    }
}
