using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLNamedRanges: IXLNamedRanges
    {
        Dictionary<String, IXLNamedRange> namedRanges = new Dictionary<String, IXLNamedRange>();
        private XLWorkbook workbook;
        public XLNamedRanges(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        #region IXLNamedRanges Members

        public IXLNamedRange NamedRange(String rangeName)
        {
            return namedRanges[rangeName];
        }

        public IXLNamedRange NamedRange(Int32 rangeIndex)
        {
            return namedRanges.ElementAt(rangeIndex).Value;
        }

        public IXLNamedRange Add(String rangeName, String rangeAddress, String comment = null)
        {
            var namedRange = new XLNamedRange(workbook, rangeName, rangeAddress, comment);
            namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public IXLNamedRange Add(String rangeName, IXLRange range, String comment = null)
        {
            var ranges = new XLRanges(range.Style);
            ranges.Add(range);
            return Add(rangeName, ranges, comment);
        }

        public IXLNamedRange Add(String rangeName, IXLRanges ranges, String comment = null)
        {
            var namedRange = new XLNamedRange(workbook, rangeName, ranges, comment);
            namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public void Delete(String rangeName)
        {
            namedRanges.Remove(rangeName);
        }

        public void Delete(Int32 rangeIndex)
        {
            namedRanges.Remove(namedRanges.ElementAt(rangeIndex).Key);
        }
        
        #endregion

        #region IEnumerable<IXLNamedRange> Members

        public IEnumerator<IXLNamedRange> GetEnumerator()
        {
            return namedRanges.Values.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion

    }
}
