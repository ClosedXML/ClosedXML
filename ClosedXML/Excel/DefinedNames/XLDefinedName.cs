#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLDefinedName : IXLDefinedName
    {
        private readonly XLDefinedNames _namedRanges;
        private String _name;

        internal XLWorkbook Workbook => _namedRanges.Workbook;

        internal XLDefinedName(XLDefinedNames container, String name, Boolean validateName, String formula, String comment)
        {
            if (validateName)
            {
                if (!XLHelper.ValidateName("named range", name, out var error))
                    throw new ArgumentException(error, nameof(name));
            }

            _namedRanges = container;
            _name = name;
            Visible = true;
            Comment = comment;

            //TODO range.Split(',') may produce incorrect result if a worksheet name contains comma. Refactoring needed.
            formula.Split(',').ForEach(r => RangeList.Add(r));
        }

        /// <inheritdoc />
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
                if (XLHelper.NameComparer.Equals(_name, value))
                    return;

                if (!XLHelper.ValidateName("named range", value, out var error))
                    throw new ArgumentException(error, nameof(value));

                if (_namedRanges.Contains(value))
                    throw new InvalidOperationException($"There is already a name '{value}'.");

                _namedRanges.Delete(_name);
                _name = value;
                _namedRanges.Add(_name, this);
            }
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

        public XLNamedRangeScope Scope => _namedRanges.Scope;

        internal List<String> RangeList { get; set; } = new();

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

        public IXLDefinedName CopyTo(IXLWorksheet targetSheet)
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

            return targetSheet.DefinedNames.Add(Name, ranges);
        }

        public IXLDefinedName SetRefersTo(String range)
        {
            RefersTo = range;
            return this;
        }

        public IXLDefinedName SetRefersTo(IXLRangeBase range)
        {
            RangeList.Clear();
            RangeList.Add(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            return this;
        }

        public IXLDefinedName SetRefersTo(IXLRanges ranges)
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
