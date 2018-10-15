using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLNamedRanges : IXLNamedRanges
    {
        private readonly Dictionary<String, IXLNamedRange> _namedRanges = new Dictionary<String, IXLNamedRange>(StringComparer.OrdinalIgnoreCase);

        internal XLWorkbook Workbook { get; set; }

        internal XLWorksheet Worksheet { get; set; }

        internal XLNamedRangeScope Scope { get; }

        public XLNamedRanges(XLWorksheet worksheet)
            : this(worksheet.Workbook)
        {
            Worksheet = worksheet;
            Scope = XLNamedRangeScope.Worksheet;
        }

        public XLNamedRanges(XLWorkbook workbook)
        {
            Workbook = workbook;
            Scope = XLNamedRangeScope.Workbook;
        }

        #region IXLNamedRanges Members

        public IXLNamedRange NamedRange(String rangeName)
        {
            if (_namedRanges.TryGetValue(rangeName, out IXLNamedRange range))
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

        public IXLNamedRange Add(String rangeName, String rangeAddress, String comment)
        {
            return Add(rangeName, rangeAddress, comment, validateName: true, validateRangeAddress: true);
        }

        /// <summary>
        /// Adds the specified range name.
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="rangeAddress">The range address.</param>
        /// <param name="comment">The comment.</param>
        /// <param name="validateName">if set to <c>true</c> validates the name.</param>
        /// <param name="validateRangeAddress">if set to <c>true</c> range address will be checked for validity.</param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException"></exception>
        /// <exception cref="ArgumentException">
        /// For named ranges in the workbook scope, specify the sheet name in the reference.
        /// </exception>
        internal IXLNamedRange Add(String rangeName, String rangeAddress, String comment, Boolean validateName, Boolean validateRangeAddress)
        {
            // When loading named ranges from an existing file, we do not validate the range address or name.
            if (validateRangeAddress)
            {
                var match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress);

                if (!match.Success)
                {
                    if (XLHelper.IsValidRangeAddress(rangeAddress))
                    {
                        IXLRange range = null;
                        if (Scope == XLNamedRangeScope.Worksheet)
                            range = Worksheet.Range(rangeAddress);
                        else if (Scope == XLNamedRangeScope.Workbook)
                            range = Workbook.Range(rangeAddress);
                        else
                            throw new NotSupportedException($"Scope {Scope} is not supported");

                        if (range == null)
                            throw new ArgumentException(string.Format(
                                "The range address '{0}' for the named range '{1}' is not a valid range.", rangeAddress,
                                rangeName));

                        if (Scope == XLNamedRangeScope.Workbook || !XLHelper.NamedRangeReferenceRegex.Match(range.ToString()).Success)
                            throw new ArgumentException(
                                "For named ranges in the workbook scope, specify the sheet name in the reference.");

                        rangeAddress = Worksheet.Range(rangeAddress).ToString();
                    }
                }
            }

            var namedRange = new XLNamedRange(this, rangeName, validateName, rangeAddress, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public IXLNamedRange Add(String rangeName, IXLRange range, String comment)
        {
            var ranges = new XLRanges { range };
            return Add(rangeName, ranges, comment);
        }

        public IXLNamedRange Add(String rangeName, IXLRanges ranges, String comment)
        {
            var namedRange = new XLNamedRange(this, rangeName, ranges, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public IXLNamedRange Add(String rangeName, IXLNamedRange namedRange)
        {
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

        /// <summary>
        /// Returns a subset of named ranges that do not have invalid references.
        /// </summary>
        public IEnumerable<IXLNamedRange> ValidNamedRanges()
        {
            return this.Where(nr => nr.IsValid);
        }

        /// <summary>
        /// Returns a subset of named ranges that do have invalid references.
        /// </summary>
        public IEnumerable<IXLNamedRange> InvalidNamedRanges()
        {
            return this.Where(nr => !nr.IsValid);
        }

        #endregion IXLNamedRanges Members

        #region IEnumerable<IXLNamedRange> Members

        public IEnumerator<IXLNamedRange> GetEnumerator()
        {
            return _namedRanges.Values.GetEnumerator();
        }

        #endregion IEnumerable<IXLNamedRange> Members

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion IEnumerable Members

        public Boolean TryGetValue(String name, out IXLNamedRange range)
        {
            if (_namedRanges.TryGetValue(name, out range)) return true;

            if (Scope == XLNamedRangeScope.Workbook)
                range = Workbook.NamedRange(name);

            return range != null;
        }

        public Boolean Contains(String name)
        {
            if (_namedRanges.ContainsKey(name)) return true;

            if (Scope == XLNamedRangeScope.Workbook)
                return Workbook.NamedRange(name) != null;
            else
                return false;
        }

        internal void OnWorksheetDeleted(string worksheetName)
        {
            _namedRanges.Values
                .Cast<XLNamedRange>()
                .ForEach(nr => nr.OnWorksheetDeleted(worksheetName));
        }
    }
}
