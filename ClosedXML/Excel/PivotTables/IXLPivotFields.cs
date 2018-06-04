// Keep this file CodeMaid organised and cleaned
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

        Boolean Contains(IXLPivotField pivotField);

        IXLPivotField Get(String sourceName);

        IXLPivotField Get(Int32 index);

        Int32 IndexOf(String sourceName);
        Int32 IndexOf(IXLPivotField pf);

        void Remove(String sourceName);
    }
}
