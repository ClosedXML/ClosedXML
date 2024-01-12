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
        private readonly Dictionary<String, XLDefinedName> _namedRanges = new(XLHelper.NameComparer);

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

        IXLDefinedName IXLNamedRanges.NamedRange(String name)
        {
            return NamedRange(name) ?? throw new ArgumentException($"Range '{name}' not found.");
        }

        internal XLDefinedName? NamedRange(String name)
        {
            if (_namedRanges.TryGetValue(name, out XLDefinedName range))
                return range;

            return null;
        }

        public IXLDefinedName Add(String name, String rangeAddress)
        {
            return Add(name, rangeAddress, null);
        }

        public IXLDefinedName Add(String name, IXLRange range)
        {
            return Add(name, range, null);
        }

        public IXLDefinedName Add(String name, IXLRanges ranges)
        {
            return Add(name, ranges, null);
        }

        public IXLDefinedName Add(String name, String rangeAddress, String? comment)
        {
            return Add(name, rangeAddress, comment, validateName: true, validateRangeAddress: true);
        }

        /// <summary>
        /// Adds the specified range name.
        /// </summary>
        /// <param name="name">Name of the range.</param>
        /// <param name="rangeAddress">The range address.</param>
        /// <param name="comment">The comment.</param>
        /// <param name="validateName">if set to <c>true</c> validates the name.</param>
        /// <param name="validateRangeAddress">if set to <c>true</c> range address will be checked for validity.</param>
        /// <exception cref="NotSupportedException"></exception>
        /// <exception cref="ArgumentException">
        /// For named ranges in the workbook scope, specify the sheet name in the reference.
        /// </exception>
        internal IXLDefinedName Add(String name, String rangeAddress, String? comment, Boolean validateName, Boolean validateRangeAddress)
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
                                name));

                        if (Scope == XLNamedRangeScope.Workbook || !XLHelper.NamedRangeReferenceRegex.Match(range.ToString()).Success)
                            throw new ArgumentException(
                                "For named ranges in the workbook scope, specify the sheet name in the reference.");

                        rangeAddress = range.ToString();
                    }
                }
            }

            var namedRange = new XLDefinedName(this, name, validateName, rangeAddress, comment);
            _namedRanges.Add(name, namedRange);
            return namedRange;
        }

        public IXLDefinedName Add(String name, IXLRange range, String? comment)
        {
            var ranges = new XLRanges { range };
            return Add(name, ranges, comment);
        }

        public IXLDefinedName Add(String name, IXLRanges ranges, String? comment)
        {
            var namedRange = new XLDefinedName(this, name, ranges, comment);
            _namedRanges.Add(name, namedRange);
            return namedRange;
        }

        internal XLDefinedName Add(String name, XLDefinedName namedRange)
        {
            _namedRanges.Add(name, namedRange);
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
        public IEnumerable<IXLDefinedName> ValidNamedRanges()
        {
            return this.Where(nr => nr.IsValid);
        }

        /// <summary>
        /// Returns a subset of named ranges that do have invalid references.
        /// </summary>
        public IEnumerable<IXLDefinedName> InvalidNamedRanges()
        {
            return this.Where(nr => !nr.IsValid);
        }

        #endregion IXLNamedRanges Members

        #region IEnumerable<IXLNamedRange> Members

        public IEnumerator<IXLDefinedName> GetEnumerator()
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

        public Boolean TryGetValue(String name, [NotNullWhen(true)] out IXLDefinedName? range)
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
