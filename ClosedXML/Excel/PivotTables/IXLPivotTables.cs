using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotTables : IEnumerable<IXLPivotTable>
    {
        IXLPivotTable Add(String name, IXLCell targetCell, IXLRange range);

        IXLPivotTable Add(String name, IXLCell targetCell, IXLTable table);

        Boolean Contains(String name);

        void Delete(String name);

        void DeleteAll();

        IXLPivotTable PivotTable(String name);
    }
}
