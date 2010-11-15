using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRanges: IEnumerable<IXLRange>, IXLStylized
    {
        void Clear();
        void Add(IXLRange range);
        void Remove(IXLRange range);
    }
}
