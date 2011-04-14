using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLNamedRange: IXLNamedRange
    {
        private List<String> rangeList = new List<String>();
        private XLNamedRanges namedRanges;
        public XLNamedRange(XLNamedRanges namedRanges , String rangeName, String range,  String comment = null)
        {
            Name = rangeName;
            rangeList.Add(range);
            Comment = comment;
            this.namedRanges = namedRanges;
        }

        public XLNamedRange(XLNamedRanges namedRanges, String rangeName, IXLRanges ranges, String comment = null)
        {
            Name = rangeName;
            ranges.ForEach(r => rangeList.Add(r.ToStringFixed()));
            Comment = comment;
            this.namedRanges = namedRanges;
        }

        public String Name { get; set; }
        public IXLRanges Ranges
        {
            get
            {
                var ranges = new XLRanges(namedRanges.Workbook, namedRanges.Workbook.Style);
                foreach (var rangeAddress in rangeList)
                {
                    var byExclamation = rangeAddress.Split('!');
                    var wsName = byExclamation[0].Replace("'", "");
                    var rng = byExclamation[1];
                    var rangeToAdd = namedRanges.Workbook.Worksheets.Worksheet(wsName).Range(rng);
                    ranges.Add(rangeToAdd);
                }
                return ranges;
            }
        }
        public IXLRange Range 
        {
            get
            {
                return Ranges.Single();
            }
        }
        public String Comment { get; set; }

        public IXLRanges Add(String rangeAddress)
        {
            var ranges = new XLRanges(namedRanges.Workbook, namedRanges.Workbook.Style);
            ranges.Add(rangeAddress);
            return Add(ranges);
        }
        public IXLRanges Add(IXLRange range)
        {
            var ranges = new XLRanges(((XLRange)range).Worksheet.Internals.Workbook, range.Style);
            ranges.Add(range);
            return Add(ranges);
        }
        public IXLRanges Add(IXLRanges ranges)
        {
            ranges.ForEach(r => rangeList.Add(r.ToString()));
            return ranges;
        }

        public void Delete()
        {
            namedRanges.Delete(Name);
        }
        public void Clear()
        {
            rangeList.Clear();
        }
        public void Remove(String rangeAddress)
        {
            rangeList.Remove(rangeAddress);
        }
        public void Remove(IXLRange range)
        {
            rangeList.Remove(range.ToString());
        }
        public void Remove(IXLRanges ranges)
        {
            ranges.ForEach(r => rangeList.Remove(r.ToString()));
        }


        public override string ToString()
        {
            String retVal = rangeList.Aggregate(String.Empty, (agg, r) => agg += r + ",");
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }
    }
}
