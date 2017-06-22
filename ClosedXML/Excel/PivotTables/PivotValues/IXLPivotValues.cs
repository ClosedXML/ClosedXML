using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPivotValues: IEnumerable<IXLPivotValue>
    {
        IXLPivotValue Add(String sourceName);
        IXLPivotValue Add(String sourceName, String customName);
        void Clear();
        void Remove(String sourceName);
    }
}
