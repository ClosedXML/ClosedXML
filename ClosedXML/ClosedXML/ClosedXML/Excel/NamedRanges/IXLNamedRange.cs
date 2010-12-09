using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLNamedRange
    {
        String Name { get; set; }
        IXLRanges Ranges { get; }
        IXLRange Range { get; }
        String Comment { get; set; }
        IXLRanges Add(String rangeAddress);
        IXLRanges Add(IXLRange range);
        IXLRanges Add(IXLRanges ranges);
        void Delete();
        void Clear();
        void Remove(String rangeAddress);
        void Remove(IXLRange range);
        void Remove(IXLRanges ranges);
    }
}
