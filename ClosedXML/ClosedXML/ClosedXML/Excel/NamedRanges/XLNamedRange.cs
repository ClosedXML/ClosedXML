using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLNamedRange: IXLNamedRange
    {
        private List<String> rangeList = new List<String>();
        private XLWorkbook workbook;
        public XLNamedRange(XLWorkbook workbook, String rangeName, String range, String comment = null)
        {
            Name = rangeName;
            rangeList.Add(range);
            Comment = comment;
            this.workbook = workbook;
        }

        public XLNamedRange(XLWorkbook workbook, String rangeName, IXLRanges ranges, String comment = null)
        {
            Name = rangeName;
            ranges.ForEach(r => rangeList.Add(r.ToString()));
            Comment = comment;
            this.workbook = workbook;
        }

        public String Name { get; set; }
        public IXLRanges Ranges
        {
            get
            {
                var ranges = new XLRanges(workbook.Style);
                foreach (var rangeAddress in rangeList)
                {
                    var byExclamation = rangeAddress.Split('!');
                    var wsName = byExclamation[0].Replace("'", "");
                    var rng = byExclamation[1];
                    var rangeToAdd = workbook.Worksheets.Worksheet(wsName).Range(rng);
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

        public override string ToString()
        {
            String retVal = rangeList.Aggregate(String.Empty, (agg, r) => agg += r + ",");
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }
    }
}
