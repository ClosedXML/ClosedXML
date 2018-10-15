using ClosedXML.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLNamedRange : IXLNamedRange
    {
        private String _name;
        private readonly XLNamedRanges _namedRanges;

        internal XLWorkbook Workbook => _namedRanges.Workbook;

        public XLNamedRange(XLNamedRanges namedRanges, String rangeName, String range, String comment = null)
            : this(namedRanges, rangeName, validateName: true, range, comment)
        {
        }

        public XLNamedRange(XLNamedRanges namedRanges, String rangeName, IXLRanges ranges, String comment = null)
            : this(namedRanges, rangeName, validateName: true, comment)
        {
            ranges.ForEach(r => RangeList.Add(r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true)));
        }

        internal XLNamedRange(XLNamedRanges namedRanges, String rangeName, Boolean validateName, String range, String comment)
            : this(namedRanges, rangeName, validateName, comment)
        {
            //TODO range.Split(',') may produce incorrect result if a worksheet name contains comma. Refactoring needed.
            range.Split(',').ForEach(r => RangeList.Add(r));
        }

        internal XLNamedRange(XLNamedRanges namedRanges, String rangeName, Boolean validateName, String comment)
        {
            _namedRanges = namedRanges ?? throw new ArgumentNullException(nameof(namedRanges));
            Visible = true;

            if (validateName)
                Name = rangeName;
            else
                SetNameWithoutValidation(rangeName);

            Comment = comment;
        }

        /// <summary>
        /// Checks if the named range contains invalid references (#REF!).
        /// </summary>
        public bool IsValid
        {
            get
            {
                return RangeList.SelectMany(c => c.Split(',')).All(r =>
                    !r.StartsWith("#REF!", StringComparison.OrdinalIgnoreCase) &&
                    !r.EndsWith("#REF!", StringComparison.OrdinalIgnoreCase));
            }
        }

        public String Name
        {
            get { return _name; }
            set
            {
                if (_name == value) return;

                var oldname = _name ?? string.Empty;

                var existingNames = _namedRanges.Select(nr => nr.Name).ToList();
                if (_namedRanges.Scope == XLNamedRangeScope.Workbook)
                    existingNames.AddRange(_namedRanges.Workbook.NamedRanges.Select(nr => nr.Name));

                if (_namedRanges.Scope == XLNamedRangeScope.Worksheet)
                    existingNames.AddRange(_namedRanges.Worksheet.NamedRanges.Select(nr => nr.Name));

                existingNames = existingNames.Distinct().ToList();

                if (!XLHelper.ValidateName("named range", value, oldname, existingNames, out String message))
                    throw new ArgumentException(message, nameof(value));

                _name = value;

                if (!String.IsNullOrWhiteSpace(oldname) && !String.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
                {
                    _namedRanges.Delete(oldname);
                    _namedRanges.Add(_name, this);
                }
            }
        }

        private void SetNameWithoutValidation(string value)
        {
            _name = value;
        }

        public IXLRanges Ranges
        {
            get
            {
                var ranges = new XLRanges();
                foreach (var rangeToAdd in
                   from rangeAddress in RangeList.SelectMany(c => c.Split(',')).Where(s => s[0] != '"')
                   let match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress)
                   select
                       match.Groups["Sheet"].Success && Workbook.Worksheets.Contains(match.Groups["Sheet"].Value)
                       ? Workbook.WorksheetsInternal.Worksheet(match.Groups["Sheet"].Value).Range(match.Groups["Range"].Value) as IXLRangeBase
                       : Workbook.Worksheets.SelectMany(sheet => sheet.Tables).SingleOrDefault(table => table.Name == match.Groups["Table"].Value)?
                               .DataRange?.Column(match.Groups["Column"].Value))
                {
                    if (rangeToAdd != null)
                        ranges.Add(rangeToAdd);
                }
                return ranges;
            }
        }

        public String Comment { get; set; }

        public Boolean Visible { get; set; }

        public XLNamedRangeScope Scope { get { return _namedRanges.Scope; } }

        public IXLRanges Add(XLWorkbook workbook, String rangeAddress)
        {
            var ranges = new XLRanges();
            var byExclamation = rangeAddress.Split('!');
            var wsName = byExclamation[0].Replace("'", "");
            var rng = byExclamation[1];
            var rangeToAdd = workbook.WorksheetsInternal.Worksheet(wsName).Range(rng);

            ranges.Add(rangeToAdd);
            return Add(ranges);
        }

        public IXLRanges Add(IXLRange range)
        {
            var ranges = new XLRanges { range };
            return Add(ranges);
        }

        public IXLRanges Add(IXLRanges ranges)
        {
            ranges.ForEach(r => RangeList.Add(r.ToString()));
            return ranges;
        }

        public void Delete()
        {
            _namedRanges.Delete(Name);
        }

        public void Clear()
        {
            RangeList.Clear();
        }

        public void Remove(String rangeAddress)
        {
            RangeList.Remove(rangeAddress);
        }

        public void Remove(IXLRange range)
        {
            RangeList.Remove(range.ToString());
        }

        public void Remove(IXLRanges ranges)
        {
            ranges.ForEach(r => RangeList.Remove(r.ToString()));
        }

        public override string ToString()
        {
            String retVal = RangeList.Aggregate(String.Empty, (agg, r) => agg + (r + ","));
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }

        public String RefersTo
        {
            get { return ToString(); }
            set
            {
                RangeList.Clear();
                RangeList.Add(value);
            }
        }

        public IXLNamedRange CopyTo(IXLWorksheet targetSheet)
        {
            if (targetSheet == _namedRanges.Worksheet)
                throw new InvalidOperationException("Cannot copy named range to the worksheet it already belongs to.");

            var ranges = new XLRanges();
            foreach (var r in Ranges)
            {
                if (_namedRanges.Worksheet == r.Worksheet)
                    // Named ranges on the source worksheet have to point to the new destination sheet
                    ranges.Add(targetSheet.Range(((XLRangeAddress)r.RangeAddress).WithoutWorksheet()));
                else
                    ranges.Add(r);
            }

            return targetSheet.NamedRanges.Add(Name, ranges);
        }

        internal IList<String> RangeList { get; set; } = new List<String>();

        public IXLNamedRange SetRefersTo(String range)
        {
            RefersTo = range;
            return this;
        }

        public IXLNamedRange SetRefersTo(IXLRangeBase range)
        {
            RangeList.Clear();
            RangeList.Add(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            return this;
        }

        public IXLNamedRange SetRefersTo(IXLRanges ranges)
        {
            RangeList.Clear();
            ranges.ForEach(r => RangeList.Add(r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true)));
            return this;
        }

        internal void OnWorksheetDeleted(string worksheetName)
        {
            var escapedSheetName = worksheetName.EscapeSheetName();
            RangeList = RangeList
                .Select(
                    rl => string.Join(",", rl
                        .Split(',')
                        .Select(r => r.StartsWith(escapedSheetName + "!", StringComparison.OrdinalIgnoreCase)
                                     ? "#REF!" + r.Substring(escapedSheetName.Length + 1)
                                     : r))
                ).ToList();
        }
    }
}
