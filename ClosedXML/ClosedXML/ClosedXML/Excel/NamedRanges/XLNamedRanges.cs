using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLNamedRanges: IXLNamedRanges
    {
        readonly Dictionary<String, IXLNamedRange> _namedRanges = new Dictionary<String, IXLNamedRange>();
        internal XLWorkbook Workbook { get; set; }
       
        public XLNamedRanges(XLWorkbook workbook)
        {
            Workbook = workbook;
        }

        #region IXLNamedRanges Members

        public IXLNamedRange NamedRange(String rangeName)
        {
            IXLNamedRange range;
            if (_namedRanges.TryGetValue(rangeName, out range))
                return range;

            return null;
        }

        public IXLNamedRange Add(String rangeName, String rangeAddress)
        {
            return Add(rangeName, rangeAddress, null);
        }
        public IXLNamedRange Add(String rangeName, IXLRange range)
        {
            return Add(rangeName, range, null);
        }
        public IXLNamedRange Add(String rangeName, IXLRanges ranges)
        {
            return Add(rangeName, ranges, null);
        }
        public IXLNamedRange Add(String rangeName, String rangeAddress, String comment )
        {
            var namedRange = new XLNamedRange(this, rangeName, rangeAddress, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }
        public IXLNamedRange Add(String rangeName, IXLRange range, String comment)
        {
            var ranges = new XLRanges {range};
            return Add(rangeName, ranges, comment);
        }
        public IXLNamedRange Add(String rangeName, IXLRanges ranges, String comment)
        {
            var namedRange = new XLNamedRange(this, rangeName, ranges, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public void Delete(String rangeName)
        {
            _namedRanges.Remove(rangeName);
        }
        public void Delete(Int32 rangeIndex)
        {
            _namedRanges.Remove(_namedRanges.ElementAt(rangeIndex).Key);
        }
        public void DeleteAll()
        {
            _namedRanges.Clear();
        }
        
        #endregion

        #region IEnumerable<IXLNamedRange> Members

        public IEnumerator<IXLNamedRange> GetEnumerator()
        {
            return _namedRanges.Values.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion

        public Boolean TryGetValue(String name, out IXLNamedRange range)
        {
            if (_namedRanges.TryGetValue(name, out range)) return true;

            range = Workbook.NamedRange(name);
            return range != null;
        }

        public Boolean Contains(String name)
        {
            if (_namedRanges.ContainsKey(name)) return true;
            return Workbook.NamedRange(name) != null;
        }

    }
}
