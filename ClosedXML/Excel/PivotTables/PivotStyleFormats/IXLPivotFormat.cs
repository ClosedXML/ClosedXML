using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    public interface IXLPivotFormat
    {
        IXLStyle Style { get; set; }
        XLPivotStyleFormatElement AppliesTo { get; set; }
        bool Outline { get; set; }
        string FieldName { get; set; }
        int? FieldIndex { get; set; }
        XLPivotTableAxisValues? Axis { get; }
        XLPivotAreaValues AreaType { get; }
        bool GrandRow { get; }
        bool GrandCol { get; }
        bool CollapsedLevelsAreSubtotals { get; }
        IEnumerable<IFieldRef> FieldReferences { get; }
    }

    public interface IXLPivotFormatList : IEnumerable<IXLPivotFormat>
    {
        IXLPivotFormat Add(Action<PivotFormatFactoryProvider> config);
        void Add(IXLPivotFormat format);
    }

    public interface IFieldRef
    {
        string FieldName { get; }
        int[] Values { get; set; }
        bool DefaultSubtotal { get; }
        bool SumSubtotal { get; }
        bool CountSubtotal { get; }
        bool CountASubtotal { get; }
        bool AverageSubtotal { get; }
        bool MaxSubtotal { get; }
        bool MinSubtotal { get; }
        bool ApplyProductInSubtotal { get; }
        bool ApplyVarianceInSubtotal { get; }
        bool ApplyVariancePInSubtotal { get; }
        bool ApplyStandardDeviationInSubtotal { get; }
        bool ApplyStandardDeviationPInSubtotal { get; }
    }
}
