using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    //internal delegate void ColumnDeletingDelegate(Int32 deletedColumn, Boolean beingShifted);
    //internal delegate void ColumnShiftedDelegate(Int32 startingColumn, Int32 columnsShifted);
    internal class XLColumnsCollection : IDictionary<Int32, XLColumn>
    {
        //public event ColumnDeletingDelegate ColumnDeleting;
        //public event ColumnShiftedDelegate ColumnShifted;

        //private Boolean beingShifted = false;
        public void ShiftColumnsRight(Int32 startingColumn, Int32 columnsToShift)
        {
            //beingShifted = true;
            foreach (var ro in dictionary.Keys.Where(k => k >= startingColumn).OrderByDescending(k => k))
            {
                var columnToMove = dictionary[ro];
                var newColumn = ro + columnsToShift;
                if (newColumn <= XLWorksheet.MaxNumberOfColumns)
                {
                    var xlColumnParameters = new XLColumnParameters(columnToMove.Worksheet, columnToMove.Style, false);
                    dictionary.Add(newColumn, new XLColumn(newColumn, xlColumnParameters));
                }
                dictionary.Remove(ro);
            }

            //if (ColumnShifted != null)
            //    ColumnShifted(startingColumn, columnsToShift);

            //beingShifted = false;
        }

        private Dictionary<Int32, XLColumn> dictionary = new Dictionary<Int32, XLColumn>();

        public void Add(int key, XLColumn value)
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
            //if (ColumnDeleting != null)
            //    ColumnDeleting(key, beingShifted);

            return dictionary.Remove(key);
        }

        public bool TryGetValue(int key, out XLColumn value)
        {
            return dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLColumn> Values
        {
            get { return dictionary.Values; }
        }

        public XLColumn this[int key]
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

        public void Add(KeyValuePair<int, XLColumn> item)
        {
            dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            //if (ColumnDeleting != null)
            //    dictionary.ForEach(r => ColumnDeleting(r.Key, beingShifted));

            dictionary.Clear();
        }

        public bool Contains(KeyValuePair<int, XLColumn> item)
        {
            return dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, XLColumn>[] array, int arrayIndex)
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

        public bool Remove(KeyValuePair<int, XLColumn> item)
        {
            //if (ColumnDeleting != null)
            //    ColumnDeleting(item.Key, beingShifted);

            return dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<int, XLColumn>> GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }
    }
}
