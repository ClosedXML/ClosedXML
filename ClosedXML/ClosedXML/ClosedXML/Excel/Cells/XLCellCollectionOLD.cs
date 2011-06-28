using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCellCollectionOLD : IDictionary<IXLAddress, XLCell>
    {
        #region Private fields
        private readonly Dictionary<IXLAddress, XLCell> m_dictionary = new Dictionary<IXLAddress, XLCell>();
        private readonly Dictionary<IXLAddress, XLCell> m_deleted = new Dictionary<IXLAddress, XLCell>();
        #endregion
        public Dictionary<IXLAddress, XLCell> Deleted
        {
            get { return m_deleted; }
        }

        public void Add(IXLAddress key, XLCell value)
        {
            if (m_deleted.ContainsKey(key))
            {
                m_deleted.Remove(key);
            }

            m_dictionary.Add(key, value);
        }

        public bool ContainsKey(IXLAddress key)
        {
            return m_dictionary.ContainsKey(key);
        }

        public ICollection<IXLAddress> Keys
        {
            get { return m_dictionary.Keys; }
        }

        public bool Remove(IXLAddress key)
        {
            if (!m_deleted.ContainsKey(key))
            {
                m_deleted.Add(key, m_dictionary[key]);
            }

            return m_dictionary.Remove(key);
        }

        public bool TryGetValue(IXLAddress key, out XLCell value)
        {
            return m_dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLCell> Values
        {
            get { return m_dictionary.Values; }
        }

        public XLCell this[IXLAddress key]
        {
            get { return m_dictionary[key]; }
            set { m_dictionary[key] = value; }
        }

        public void Add(KeyValuePair<IXLAddress, XLCell> item)
        {
            if (m_deleted.ContainsKey(item.Key))
            {
                m_deleted.Remove(item.Key);
            }
            m_dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            foreach (var kp in m_dictionary)
            {
                if (!m_deleted.ContainsKey(kp.Key))
                {
                    m_deleted.Add(kp.Key, kp.Value);
                }
            }
            m_dictionary.Clear();
        }

        public bool Contains(KeyValuePair<IXLAddress, XLCell> item)
        {
            return m_dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<IXLAddress, XLCell>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return m_dictionary.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(KeyValuePair<IXLAddress, XLCell> item)
        {
            if (!m_deleted.ContainsKey(item.Key))
            {
                m_deleted.Add(item.Key, m_dictionary[item.Key]);
            }

            return m_dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<IXLAddress, XLCell>> GetEnumerator()
        {
            return m_dictionary.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return m_dictionary.GetEnumerator();
        }

        public void RemoveAll()
        {
            RemoveAll(c => true);
        }

        public void RemoveAll(Func<XLCell, Boolean> predicate)
        {
            foreach (var kp in m_dictionary.Values.Where(predicate).Select(c => c))
            {
                if (!m_deleted.ContainsKey(kp.Address))
                {
                    m_deleted.Add(kp.Address, kp);
                }
            }

            m_dictionary.RemoveAll(predicate);
        }
    }
}