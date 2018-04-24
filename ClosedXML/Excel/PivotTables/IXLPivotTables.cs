using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotTables : IEnumerable<IXLPivotTable>
    {
        void Add(String name, IXLPivotTable pivotTable);

        IXLPivotTable AddNew(String name, IXLCell target, IXLRange source);

        Boolean Contains(String name);

        void Delete(String name);

        void DeleteAll();

        IXLPivotTable PivotTable(String name);
    }
}
