using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotElementStyleFormats
    {
        IXLPivotStyleFormat Label { get; }
        IXLPivotValueStyleFormat AddValuesFormat();
        IEnumerable<IXLPivotValueStyleFormat> DataValuesFormats { get; }
        bool HasLabelFormat { get; }
    }
}
