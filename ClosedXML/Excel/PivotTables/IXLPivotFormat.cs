using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    public interface IXLPivotFormat
    {
        IXLStyle Style { get; }
        bool DataOnly { get; }
        bool LabelOnly { get; }
        bool Outline { get; }
        string FieldName { get; }
        PivotTableAxisValues? Axis { get; }
        PivotAreaValues AreaType { get; }
        bool GrandRow { get; }
        bool GrandCol { get; }
        bool CollapsedLevelsAreSubtotals { get; }
        IEnumerable<IFieldRef> FieldReferences { get; }
    }

    public interface IXLPivotFormatList : IEnumerable<IXLPivotFormat>
    {
        IXLPivotFormat Add(Action<PivotFormatFactoryProvider> config);
    }

    public interface IFieldRef
    {
        string FieldName { get; }
        object Value { get; }
        bool DefaultSubtotal { get; }
    }
}
