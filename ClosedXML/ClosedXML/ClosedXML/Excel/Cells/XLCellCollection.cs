using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCellCollection : IDictionary<IXLAddress, XLCell>
    {
        private Dictionary<IXLAddress, XLCell> dictionary = new Dictionary<IXLAddress, XLCell>();

        private Dictionary<IXLAddress, XLCell> deleted = new Dictionary<IXLAddress, XLCell>();
        public Dictionary<IXLAddress, XLCell> Deleted
        {
            get
            {
                return deleted;
            }
        }

        public void Add(IXLAddress key, XLCell value)
        {
            if (deleted.ContainsKey(key))
                deleted.Remove(key);

            dictionary.Add(key, value);
        }

        public bool ContainsKey(IXLAddress key)
        {
            return dictionary.ContainsKey(key);
        }

        public ICollection<IXLAddress> Keys
        {
            get { return dictionary.Keys; }
        }

        public bool Remove(IXLAddress key)
        {
            if (!deleted.ContainsKey(key))
                deleted.Add(key, dictionary[key]);

            return dictionary.Remove(key);
        }

        public bool TryGetValue(IXLAddress key, out XLCell value)
        {
            return dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLCell> Values
        {
            get { return dictionary.Values; }
        }

        public XLCell this[IXLAddress key]
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

        public void Add(KeyValuePair<IXLAddress, XLCell> item)
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

        public bool Contains(KeyValuePair<IXLAddress, XLCell> item)
        {
            return dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<IXLAddress, XLCell>[] array, int arrayIndex)
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

        public bool Remove(KeyValuePair<IXLAddress, XLCell> item)
        {
            if (!deleted.ContainsKey(item.Key))
                deleted.Add(item.Key, dictionary[item.Key]);

            return dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<IXLAddress, XLCell>> GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        public void RemoveAll()
        {
            RemoveAll(c => true);
        }

        public void RemoveAll(Func<XLCell, Boolean> predicate)
        {
            foreach (var kp in dictionary.Values.Where(predicate).Select(c=>c))
            {
                if (!deleted.ContainsKey(kp.Address))
                    deleted.Add(kp.Address, kp);
            }

            dictionary.RemoveAll(predicate);
        }
    }
}
