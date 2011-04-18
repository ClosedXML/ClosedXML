using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    //internal delegate void RowDeletingDelegate(Int32 deletedRow, Boolean beingShifted);
    //internal delegate void RowShiftedDelegate(Int32 startingRow, Int32 rowsShifted);
    internal class XLRowsCollection: IDictionary<Int32, XLRow>
    {
        //public event RowDeletingDelegate RowDeleting;
        //public event RowShiftedDelegate RowShifted;

        //private Boolean beingShifted = false;
        public void ShiftRowsDown(Int32 startingRow, Int32 rowsToShift)
        {
            //beingShifted = true;
            foreach (var ro in dictionary.Keys.Where(k => k >= startingRow).OrderByDescending(k => k))
            {
                var rowToMove = dictionary[ro];
                Int32 newRow = ro + rowsToShift;
                if (newRow <= XLWorksheet.MaxNumberOfRows)
                {
                    dictionary.Add(newRow, new XLRow(rowToMove, rowToMove.Worksheet));
                }
                dictionary.Remove(ro);
            }

            //if (RowShifted != null)
            //    RowShifted(startingRow, rowsToShift);

            //beingShifted = false;
        }

        private Dictionary<Int32, XLRow> dictionary = new Dictionary<Int32, XLRow>();

        public void Add(int key, XLRow value)
        {
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
            //if (RowDeleting != null)
            //    RowDeleting(key, beingShifted);

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
            dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            //if (RowDeleting != null)
            //    dictionary.ForEach(r => RowDeleting(r.Key, beingShifted));

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
            //if (RowDeleting != null)
            //    RowDeleting(item.Key, beingShifted);

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
    }
}
