// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotTableCalculatedField : IXLPivotTableCalculatedField
    {
        public XLPivotTableCalculatedField(String name, String formula)
        {
            this.Name = name;
            this.Formula = formula;
        }

        public String Formula { get; set; }
        public String Name { get; set; }
    }
}
