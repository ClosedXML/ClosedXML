using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A collection of a named ranges, either for workbook or for worksheet.
    /// </summary>
    internal class XLNamedRanges : IXLNamedRanges
    {
        private readonly Dictionary<String, XLNamedRange> _namedRanges = new(XLHelper.NameComparer);

        internal XLWorkbook Workbook { get; set; }

        internal XLWorksheet? Worksheet { get; set; }

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

        IXLNamedRange? IXLNamedRanges.NamedRange(String rangeName) => NamedRange(rangeName);

        internal XLNamedRange? NamedRange(String rangeName)
        {
            if (_namedRanges.TryGetValue(rangeName, out XLNamedRange range))
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

        public IXLNamedRange Add(String rangeName, String rangeAddress, String? comment)
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
        /// <exception cref="NotSupportedException"></exception>
        /// <exception cref="ArgumentException">
        /// For named ranges in the workbook scope, specify the sheet name in the reference.
        /// </exception>
        internal IXLNamedRange Add(String rangeName, String rangeAddress, String? comment, Boolean validateName, Boolean validateRangeAddress)
        {
            // When loading named ranges from an existing file, we do not validate the range address or name.
            if (validateRangeAddress)
            {
                var match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress);

                if (!match.Success)
                {
                    if (XLHelper.IsValidRangeAddress(rangeAddress))
                    {
                        IXLRange? range;
                        if (Scope == XLNamedRangeScope.Worksheet)
                            range = Worksheet!.Range(rangeAddress);
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

                        rangeAddress = range.ToString();
                    }
                }
            }

            var namedRange = new XLNamedRange(this, rangeName, validateName, rangeAddress, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        public IXLNamedRange Add(String rangeName, IXLRange range, String? comment)
        {
            var ranges = new XLRanges { range };
            return Add(rangeName, ranges, comment);
        }

        public IXLNamedRange Add(String rangeName, IXLRanges ranges, String? comment)
        {
            var namedRange = new XLNamedRange(this, rangeName, ranges, comment);
            _namedRanges.Add(rangeName, namedRange);
            return namedRange;
        }

        internal XLNamedRange Add(String rangeName, XLNamedRange namedRange)
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

        public Boolean TryGetValue(String name, [NotNullWhen(true)] out IXLNamedRange? range)
        {
            if (_namedRanges.TryGetValue(name, out var rangeInternal))
            {
                range = rangeInternal;
                return true;
            }

            range = Scope == XLNamedRangeScope.Workbook
                ? Workbook.NamedRange(name)
                : null;

            return range is not null;
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
                .ForEach(nr => nr.OnWorksheetDeleted(worksheetName));
        }
    }
}
