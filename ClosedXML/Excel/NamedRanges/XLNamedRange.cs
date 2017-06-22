using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Text.RegularExpressions;

    internal class XLNamedRange: IXLNamedRange
    {

        private List<String> _rangeList = new List<String>();
        private readonly XLNamedRanges _namedRanges;
        public XLNamedRange(XLNamedRanges namedRanges , String rangeName, String range,  String comment = null)
        {
            Visible = true;
            Name = rangeName;
            _rangeList.Add(range);
            Comment = comment;
            _namedRanges = namedRanges;
        }

        public XLNamedRange(XLNamedRanges namedRanges, String rangeName, IXLRanges ranges, String comment = null)
        {
            Visible = true;
            Name = rangeName;
            ranges.ForEach(r => _rangeList.Add(r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true)));
            Comment = comment;
            _namedRanges = namedRanges;
        }

        public String Name { get; set; }
        public IXLRanges Ranges
        {
            get
            {
                var ranges = new XLRanges();
                foreach (var rangeToAdd in
                   from rangeAddress in _rangeList.SelectMany(c=>c.Split(',')).Where(s=>s[0] != '"')
                   let match = XLHelper.NamedRangeReferenceRegex.Match(rangeAddress)
                   select
                       match.Groups["Sheet"].Success
                       ?  _namedRanges.Workbook.WorksheetsInternal.Worksheet(match.Groups["Sheet"].Value).Range(match.Groups["Range"].Value) as IXLRangeBase
                       : _namedRanges.Workbook.Worksheets.SelectMany(sheet => sheet.Tables).Single(table => table.Name == match.Groups["Table"].Value).DataRange.Column(match.Groups["Column"].Value) )
                {
                    ranges.Add(rangeToAdd);
                }
                return ranges;
            }
        }

        public String Comment { get; set; }

        public Boolean Visible { get; set; }

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
            var ranges = new XLRanges {range};
            return Add(ranges);
        }
        public IXLRanges Add(IXLRanges ranges)
        {
            ranges.ForEach(r => _rangeList.Add(r.ToString()));
            return ranges;
        }

        public void Delete()
        {
            _namedRanges.Delete(Name);
        }
        public void Clear()
        {
            _rangeList.Clear();
        }
        public void Remove(String rangeAddress)
        {
            _rangeList.Remove(rangeAddress);
        }
        public void Remove(IXLRange range)
        {
            _rangeList.Remove(range.ToString());
        }
        public void Remove(IXLRanges ranges)
        {
            ranges.ForEach(r => _rangeList.Remove(r.ToString()));
        }


        public override string ToString()
        {
            String retVal = _rangeList.Aggregate(String.Empty, (agg, r) => agg + (r + ","));
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }

        public String RefersTo
        {
            get { return ToString(); }
            set
            {
                _rangeList.Clear();
                _rangeList.Add(value);
            }
        }

        internal List<String> RangeList
        {
            get { return _rangeList; }
            set { _rangeList = value; }
        }

        public IXLNamedRange SetRefersTo(String range)
        {
            RefersTo = range;
            return this;
        }
        public IXLNamedRange SetRefersTo(IXLRangeBase range)
        {
            _rangeList.Clear();
            _rangeList.Add(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            return this;
        }
        public IXLNamedRange SetRefersTo(IXLRanges ranges)
        {
            _rangeList.Clear();
            ranges.ForEach(r => _rangeList.Add(r.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true)));
            return this;
        }
    }
}
