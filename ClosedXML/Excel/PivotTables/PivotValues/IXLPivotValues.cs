// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotValues : IEnumerable<IXLPivotValue>
    {
        IXLPivotValue Add(string sourceName);

        IXLPivotValue Add(string sourceName, string customName);

        void Clear();

        bool Contains(string sourceName);

        bool Contains(IXLPivotValue pivotValue);

        IXLPivotValue Get(string sourceName);

        IXLPivotValue Get(int index);

        int IndexOf(string sourceName);

        int IndexOf(IXLPivotValue pivotValue);

        void Remove(string sourceName);
    }
}
