using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal abstract class AbstractPivotFieldReference: IEquatable<PivotAreaReference>
    {
        public bool Subtotal { get; set; }

        internal abstract bool DefaultSubtotal { get; }
        internal abstract bool SumSubtotal { get; }
        internal abstract bool CountSubtotal { get; }
        internal abstract bool CountASubtotal { get; }
        internal abstract bool AverageSubtotal { get; }
        internal abstract bool MaxSubtotal { get; }
        internal abstract bool MinSubtotal { get; }
        internal abstract bool ApplyProductInSubtotal { get; }
        internal abstract bool ApplyVarianceInSubtotal { get; }
        internal abstract bool ApplyVariancePInSubtotal { get; }
        internal abstract bool ApplyStandardDeviationInSubtotal { get; }
        internal abstract bool ApplyStandardDeviationPInSubtotal { get; }
        internal abstract UInt32Value GetFieldOffset();

        /// <summary>
        ///   <P>Helper function used during saving to calculate the indices of the filtered values</P>
        /// </summary>
        /// <returns>Indices of the filtered values</returns>
        internal abstract IEnumerable<Int32> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt);

        public bool Equals(PivotAreaReference other)
        {
            return !ReferenceEquals(null, other)
                   && GetFieldOffset().ToString() == other.Field.ToString()
                   && DefaultSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.DefaultSubtotal, false)
                   && SumSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.SumSubtotal, false)
                   && CountSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.CountSubtotal, false)
                   && CountASubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.CountASubtotal, false)
                   && AverageSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.AverageSubtotal, false)
                   && MaxSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.MaxSubtotal, false)
                   && MinSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.MinSubtotal, false)
                   && ApplyProductInSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.ApplyProductInSubtotal, false)
                   && ApplyVarianceInSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.ApplyVarianceInSubtotal, false)
                   && ApplyVariancePInSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.ApplyVariancePInSubtotal, false)
                   && ApplyStandardDeviationInSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.ApplyStandardDeviationInSubtotal, false)
                   && ApplyStandardDeviationPInSubtotal == OpenXmlHelper.GetBooleanValueAsBool(other.ApplyStandardDeviationPInSubtotal, false);
        }

        public static explicit operator PivotAreaReference(AbstractPivotFieldReference value)
        {
            return new PivotAreaReference
            {
                SumSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.SumSubtotal, false),
                CountASubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.CountASubtotal, false),
                AverageSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.AverageSubtotal, false),
                MaxSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.MaxSubtotal, false),
                MinSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.MinSubtotal, false),
                ApplyProductInSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.ApplyProductInSubtotal, false),
                CountSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.CountSubtotal, false),
                ApplyVarianceInSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.ApplyVarianceInSubtotal, false),
                ApplyVariancePInSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.ApplyVariancePInSubtotal, false),
                ApplyStandardDeviationInSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.ApplyStandardDeviationInSubtotal, false),
                ApplyStandardDeviationPInSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.ApplyStandardDeviationPInSubtotal, false),
                DefaultSubtotal = OpenXmlHelper.GetBooleanValue(value.Subtotal && value.DefaultSubtotal, false),
                Field = value.GetFieldOffset()
            };
        }
    }
}
