using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotTables : IEnumerable<IXLPivotTable>
    {
        IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range);

        IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table);

        [Obsolete("Use Add instead")]
        IXLPivotTable AddNew(string name, IXLCell targetCell, IXLRange range);

        [Obsolete("Use Add instead")]
        IXLPivotTable AddNew(string name, IXLCell targetCell, IXLTable table);

        bool Contains(string name);

        void Delete(string name);

        void DeleteAll();

        IXLPivotTable PivotTable(string name);
    }
}
