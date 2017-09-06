using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotFields : IEnumerable<IXLPivotField>
    {
        IXLPivotField Add(String sourceName);

        IXLPivotField Add(String sourceName, String customName);

        void Clear();

        Boolean Contains(String sourceName);

        IXLPivotField Get(String sourceName);

        Int32 IndexOf(IXLPivotField pf);

        void Remove(String sourceName);
    }
}
