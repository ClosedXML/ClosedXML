// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotValues : IEnumerable<IXLPivotValue>
    {
        IXLPivotValue Add(String sourceName);

        IXLPivotValue Add(String sourceName, String customName);

        void Clear();

        Boolean Contains(String sourceName);

        Boolean Contains(IXLPivotValue pivotValue);

        IXLPivotValue Get(String sourceName);

        IXLPivotValue Get(Int32 index);

        Int32 IndexOf(String sourceName);

        Int32 IndexOf(IXLPivotValue pivotValue);

        void Remove(String sourceName);
    }
}
