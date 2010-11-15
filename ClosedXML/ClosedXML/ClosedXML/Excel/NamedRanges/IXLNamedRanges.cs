using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLNamedRanges: IEnumerable<IXLNamedRange>
    {
        IXLNamedRange NamedRange(String rangeName);
        IXLNamedRange NamedRange(Int32 rangeIndex);
        IXLNamedRange Add(String rangeName, String rangeAddress, String comment = null);
        IXLNamedRange Add(String rangeName, IXLRange range, String comment = null);
        IXLNamedRange Add(String rangeName, IXLRanges ranges, String comment = null);
        void Delete(String rangeName);
        void Delete(Int32 rangeIndex);
    }
}
