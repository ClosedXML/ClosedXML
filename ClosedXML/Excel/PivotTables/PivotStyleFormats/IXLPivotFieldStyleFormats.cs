// Keep this file CodeMaid organised and cleaned

using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotFieldStyleFormats: IXLPivotElementStyleFormats
    {
        IXLPivotStyleFormat Header { get; }
        IXLPivotElementStyleFormats Subtotal { get; }
    }
}
