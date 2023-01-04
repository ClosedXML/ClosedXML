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

        Boolean Contains(String customName);

        Boolean Contains(IXLPivotValue pivotValue);

        IXLPivotValue Get(String customName);

        IXLPivotValue Get(Int32 index);

        Int32 IndexOf(String customName);

        Int32 IndexOf(IXLPivotValue pivotValue);

        void Remove(String customName);
    }
}
