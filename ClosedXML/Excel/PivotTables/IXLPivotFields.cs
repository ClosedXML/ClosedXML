// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotFields : IEnumerable<IXLPivotField>
    {
        IXLPivotField Add(string sourceName);

        IXLPivotField Add(string sourceName, string customName);

        void Clear();

        bool Contains(string sourceName);

        bool Contains(IXLPivotField pivotField);

        IXLPivotField Get(string sourceName);

        IXLPivotField Get(int index);

        int IndexOf(string sourceName);
        int IndexOf(IXLPivotField pf);

        void Remove(string sourceName);
    }
}
