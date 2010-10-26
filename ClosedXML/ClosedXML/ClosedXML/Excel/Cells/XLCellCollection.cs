using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    //internal delegate void RangeShiftedDelegate(XLRange range, Int32 cellsToShift, XLShiftDirection shiftDirection);
    
    internal class XLCellsCollection : IDictionary<Int32, XLCell>
    {
        //public event RangeShiftedDelegate RangeShifted;

        //public void ShiftRange(XLRange range, Int32 cellsToShift, XLShiftDirection shiftDirection)
        //{

        //    foreach (var ro in dictionary.Keys.Where(k => k >= startingCell).OrderByDescending(k => k))
        //    {
        //        var cellToMove = dictionary[ro];
        //        var newCell = ro + cellsToShift;
        //        if (newCell <= XLWorksheet.MaxNumberOfCells)
        //        {
        //            var xlCellParameters = new XLCellParameters(cellToMove.Worksheet, cellToMove.Style, false);
        //            dictionary.Add(newCell, new XLCell(newCell, xlCellParameters));
        //        }
        //        dictionary.Remove(ro);

        //        if (RangeShifted != null)
        //            RangeShifted(ro, newCell);
        //    }
        //}

        private Dictionary<Int32, XLCell> dictionary = new Dictionary<Int32, XLCell>();

        public void Add(int key, XLCell value)
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
            return dictionary.Remove(key);
        }

        public bool TryGetValue(int key, out XLCell value)
        {
            return dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLCell> Values
        {
            get { return dictionary.Values; }
        }

        public XLCell this[int key]
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

        public void Add(KeyValuePair<int, XLCell> item)
        {
            dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            dictionary.Clear();
        }

        public bool Contains(KeyValuePair<int, XLCell> item)
        {
            return dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, XLCell>[] array, int arrayIndex)
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

        public bool Remove(KeyValuePair<int, XLCell> item)
        {
            return dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<int, XLCell>> GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }
    }
}
