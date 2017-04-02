using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPivotTables: IEnumerable<IXLPivotTable>
    {
        IXLPivotTable PivotTable(String name);
        void Add(String name, IXLPivotTable pivotTable);
        IXLPivotTable AddNew(String name, IXLCell target, IXLRange source);
        void Delete(String name);
        void DeleteAll();
    }
}
