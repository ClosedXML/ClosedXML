using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotTableCalculatedField
    {
        String Formula { get; set; }
        String Name { get; }
    }
}
