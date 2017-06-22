using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPivotFields: IEnumerable<IXLPivotField>
    {
        IXLPivotField Add(String sourceName);
        IXLPivotField Add(String sourceName, String customName);
        void Clear();
        void Remove(String sourceName);

        int IndexOf(IXLPivotField pf);
    }
}
