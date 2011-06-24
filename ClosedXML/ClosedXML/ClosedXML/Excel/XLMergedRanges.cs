using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    internal sealed class XLMergedRanges:IEnumerable<SheetRange>
    {
        #region Private fields
        private readonly SortedDictionary<SheetPoint, SheetRange> m_dict = new SortedDictionary<SheetPoint, SheetRange>(SheetPointComparer.Instance);
        #endregion
        #region Constructor
        public XLMergedRanges()
        {
        }
        private XLMergedRanges(XLMergedRanges original)
        {
            m_dict = new SortedDictionary<SheetPoint, SheetRange>(original.m_dict, SheetPointComparer.Instance);
        }
        #endregion
        #region Public properties
        public int Count
        {
            [DebuggerStepThrough]
            get { return m_dict.Count; }
        }
        #endregion
        #region Public methods
        public void Add(SheetRange range)
        {
            #region Check
            if (range.IsOneCell)
            {
                throw new ArgumentException("One cell can't be merged");
            }
            #endregion
            m_dict.Add(range.FirstAddress, range);
        }

        public bool Remove(SheetPoint point)
        {
            return m_dict.Remove(point);
        }
        public bool Remove(SheetRange sheetRange)
        {
            return m_dict.Remove(sheetRange.FirstAddress);
        }

        public bool Intersects(IXLAddress address)
        {
            var point = new SheetPoint(address.RowNumber, address.ColumnNumber);
            return Intersects(new SheetRange(point, point));
        }

        public bool Intersects(SheetPoint point)
        {
            return Intersects(new SheetRange(point, point));
        }

        public bool Intersects(SheetRange range)
        {
            foreach (var mergedRange in m_dict.Values)
            {
                if (mergedRange.Intersects(range))
                {
                    return true;
                }
            }
            return false;
        }

        public void Clear()
        {
            m_dict.Clear();
        }

        public XLMergedRanges Clone()
        {
            return new XLMergedRanges(this);
        }
        #endregion
        #region Implementation of IEnumerable
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion
        #region Implementation of IEnumerable<out SheetRange>
        public IEnumerator<SheetRange> GetEnumerator()
        {
            return m_dict.Values.GetEnumerator();
        }
        #endregion
        //--
        #region Nested type: SheetPointComparer
        private sealed class SheetPointComparer : Comparer<SheetPoint>
        {
            ///<summary>
            /// Singleton instance 
            ///</summary>
            public static readonly SheetPointComparer Instance = new SheetPointComparer();
            #region Constructor
            private SheetPointComparer()
            {
            }
            #endregion
            #region Public methods
            public override int Compare(SheetPoint x, SheetPoint y)
            {
                return Math.Sign(x.RowNumber - y.RowNumber) * 2 + Math.Sign(x.ColumnNumber - y.ColumnNumber);
            }
            #endregion
            
        }

        #endregion
        
    }
}