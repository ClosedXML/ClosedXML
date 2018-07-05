using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotTableCalculatedFields : IEnumerable<IXLPivotTableCalculatedField>
    {
        IXLPivotTableCalculatedField Add(String name, String formula);

        void Clear();

        Boolean Contains(String name);

        IXLPivotTableCalculatedField Get(String name);

        void Remove(String name);

        Boolean TryGetCalculatedField(String name, out IXLPivotTableCalculatedField calculatedField);
    }
}
