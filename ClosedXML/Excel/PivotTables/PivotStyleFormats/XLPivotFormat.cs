using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLPivotFormat : IXLPivotFormat
    {
        public IXLStyle Style { get; set; }
        public XLPivotStyleFormatElement AppliesTo { get; set; } = XLPivotStyleFormatElement.Data;
        public bool Outline { get; set; } = true;
        public string FieldName { get; set; }
        public int? FieldIndex { get; set; }
        public XLPivotTableAxisValues? Axis { get; set; } = null;
        public XLPivotAreaValues AreaType { get; set; } = XLPivotAreaValues.Normal;
        public bool GrandRow { get; set; } = false;
        public bool GrandCol { get; set; } = false;
        public bool CollapsedLevelsAreSubtotals { get; set; } = false;
        public IEnumerable<IFieldRef> FieldReferences { get; } = new List<IFieldRef>();

        public XLPivotFormat(IXLStyle style = null)
        {
            Style = style ?? XLStyle.Default;
        }

        public void Build(XLPivotTable pt, XLWorkbook.SaveContext context)
        {
            if (Axis.HasValue && !string.IsNullOrWhiteSpace(FieldName))
            {
                if (pt.ColumnLabels.Contains(FieldName))
                    Axis = XLPivotTableAxisValues.AxisColumn;
                else if (pt.RowLabels.Contains(FieldName))
                    Axis = XLPivotTableAxisValues.AxisRow;
                else if (pt.ReportFilters.Contains(FieldName))
                    Axis = XLPivotTableAxisValues.AxisPage;
                else Axis = XLPivotTableAxisValues.AxisValues;
            }

            foreach (var field in FieldReferences.Cast<FieldRef>())
            {
                if (field.FieldName == XLConstants.PivotTableValuesSentinalLabel)
                {
                    field.FieldIdx = -2;
                    if (field.FieldNames != null)
                        field.Values = field.FieldNames.Select(x => ValueIndexOf(pt.Values, x)).ToArray();
                }
                else
                {
                    field.FieldIdx = FieldIndexOf(pt.Fields, field.FieldName);
                    var values = context.PivotTables[pt.Guid].Fields[field.FieldName].DistinctValues.ToList();

                    if (field.ValueFilter != null)
                    {
                        field.Values = values.Select((val, idx) => new {val, idx}).Where(v => field.ValueFilter(v.val)).Select(f => f.idx).ToArray();
                    }
                }
            }
        }

        public Int32 FieldIndexOf(IXLPivotFields fields, string fieldName)
        {
            var selectedItem = fields.Select((fld, idx) => new { fld, idx }).FirstOrDefault(i => i.fld.SourceName == fieldName);
            if (selectedItem == null)
                throw new ArgumentNullException(nameof(fieldName), $"Invalid field name {fieldName}.");

            return selectedItem.idx;
        }

        private int ValueIndexOf(IXLPivotValues values, string fieldName)
        {
            var selectedItem = values.Select((item, index) => new { Item = item, Position = index })
                .FirstOrDefault(i => i.Item.SourceName == fieldName);
            if (selectedItem == null)
                throw new ArgumentNullException(nameof(fieldName), $"Invalid field with name {fieldName}.");

            return selectedItem.Position;
        }
    }

    public class PivotFormatFactoryProvider
    {
        public IXLPivotFormat Format { get; }
        internal PivotFormatBuilder Builder { get; private set; }

        internal PivotFormatFactoryProvider(IXLStyle style)
        {
            Format = new XLPivotFormat(style);
        }

        public HeaderPivotFormatBuilder ForHeader(IXLPivotField field)
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");
            Builder = new HeaderPivotFormatBuilder(field, Format);
            return (HeaderPivotFormatBuilder)Builder;
        }

        public DataPivotFormatBuilder ForLabel(IXLPivotField field)
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");

            Builder = new DataPivotFormatBuilder(Format)
                .AndWith(field)
                .LabelOnly();
            return (DataPivotFormatBuilder)Builder;
        }

        public GrandTotalPivotFormatBuilder ForGrandRow()
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");
            Builder = new GrandTotalPivotFormatBuilder(Format, PivotTableAxisValues.AxisRow);
            return (GrandTotalPivotFormatBuilder)Builder;
        }

        public GrandTotalPivotFormatBuilder ForGrandColumn()
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");
            Builder = new GrandTotalPivotFormatBuilder(Format, PivotTableAxisValues.AxisColumn);
            return (GrandTotalPivotFormatBuilder)Builder;
        }

        public SubtotalPivotFormatBuilder ForSubtotal(IXLPivotField field)
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");
            Builder = new SubtotalPivotFormatBuilder(field, Format);
            return (SubtotalPivotFormatBuilder)Builder;
        }

        public DataPivotFormatBuilder ForData(IXLPivotField field, Predicate<object> predicate = null)
        {
            if (Builder != null)
                throw new InvalidOperationException("Builder already instantiated.");
            var builder = new DataPivotFormatBuilder(Format);
            builder.AndWith(field, predicate);
            Builder = builder;
            return builder;
        }
    }

    public abstract class PivotFormatBuilder
    {
        protected IXLPivotFormat Format;

        public IXLStyle Style => Format.Style;

        internal PivotFormatBuilder(IXLPivotFormat format)
        {
            Format = format;
        }

        internal abstract XLPivotFormat Build();
    }

    public class HeaderPivotFormatBuilder : PivotFormatBuilder
    {
        private readonly IXLPivotField _field;

        internal HeaderPivotFormatBuilder(IXLPivotField field, IXLPivotFormat format)
            : base(format)
        {
            _field = field;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat)Format;
            format.AppliesTo = XLPivotStyleFormatElement.Label;
            format.Outline = false;
            format.AreaType = XLPivotAreaValues.Button;
            format.FieldName = _field.SourceName;
            format.FieldIndex = _field.Offset;
            format.Axis = XLPivotTableAxisValues.AxisValues;
            return format;
        }
    }

    public class DataPivotFormatBuilder : PivotFormatBuilder
    {
        private XLPivotStyleFormatElement _appliesTo = XLPivotStyleFormatElement.Data;
        private readonly List<FieldRef> _refFields = new List<FieldRef>();

        internal DataPivotFormatBuilder(IXLPivotFormat format)
            : base(format)
        {
        }

        public DataPivotFormatBuilder ForValueField(IXLPivotValue valueField)
        {
            _refFields.Add(FieldRef.ValueField(valueField.SourceName));
            return this;
        }

        public DataPivotFormatBuilder AndWith(IXLPivotField field, Predicate<object> predicate = null)
        {
            _refFields.Add(FieldRef.ForField(field.SourceName, predicate));
            return this;
        }

        public DataPivotFormatBuilder DataOnly()
        {
            _appliesTo = XLPivotStyleFormatElement.Data;
            return this;
        }

        public DataPivotFormatBuilder LabelOnly()
        {
            _appliesTo = XLPivotStyleFormatElement.Label;
            return this;
        }

        public DataPivotFormatBuilder AppliesTo(XLPivotStyleFormatElement appliesTo)
        {
            _appliesTo = appliesTo;
            return this;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat)Format;

            ((List<IFieldRef>)format.FieldReferences).AddRange(_refFields);
            format.AppliesTo = _appliesTo;
            return format;
        }
    }

    public class SubtotalPivotFormatBuilder : PivotFormatBuilder
    {
        private XLPivotStyleFormatElement _appliesTo = XLPivotStyleFormatElement.All;
        private readonly IXLPivotField _field;

        internal SubtotalPivotFormatBuilder(IXLPivotField field, IXLPivotFormat format)
            : base(format)
        {
            _field = field;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat) Format;
            var fieldRef = FieldRef.ForField(_field.SourceName);
            fieldRef.DefaultSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Automatic);
            fieldRef.SumSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Sum);
            fieldRef.CountSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Count);
            fieldRef.CountASubtotal = _field.Subtotals.Contains(XLSubtotalFunction.CountNumbers);
            fieldRef.AverageSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Average);
            fieldRef.MaxSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Maximum);
            fieldRef.MinSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Minimum);
            fieldRef.ApplyProductInSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Product);
            fieldRef.ApplyVarianceInSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.Variance);
            fieldRef.ApplyVariancePInSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.PopulationVariance);
            fieldRef.ApplyStandardDeviationInSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.StandardDeviation);
            fieldRef.ApplyStandardDeviationPInSubtotal = _field.Subtotals.Contains(XLSubtotalFunction.PopulationStandardDeviation);
            ((List<IFieldRef>) format.FieldReferences).Add(fieldRef);
            format.AppliesTo = _appliesTo;
            format.Outline = false;
            return format;
        }

        public SubtotalPivotFormatBuilder AppliesTo(XLPivotStyleFormatElement appliesTo)
        {
            _appliesTo = appliesTo;
            return this;
        }
    }

    public class GrandTotalPivotFormatBuilder : PivotFormatBuilder
    {
        private readonly PivotTableAxisValues _axis;
        private XLPivotStyleFormatElement _appliesTo = XLPivotStyleFormatElement.All;

        internal GrandTotalPivotFormatBuilder(IXLPivotFormat format, PivotTableAxisValues axis) : base(format)
        {
            _axis = axis;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat)Format;

            if (_axis == PivotTableAxisValues.AxisRow)
            {
                format.GrandRow = true;
                format.Axis = XLPivotTableAxisValues.AxisRow;
            }
            else if (_axis == PivotTableAxisValues.AxisColumn)
            {
                format.GrandCol = true;
                format.Axis = XLPivotTableAxisValues.AxisColumn;
            }

            format.AppliesTo = _appliesTo;
            return format;
        }

        public GrandTotalPivotFormatBuilder DataOnly()
        {
            _appliesTo = XLPivotStyleFormatElement.Data;
            return this;
        }

        public GrandTotalPivotFormatBuilder LabelOnly()
        {
            _appliesTo = XLPivotStyleFormatElement.Label;
            return this;
        }

        public GrandTotalPivotFormatBuilder AppliesTo(XLPivotStyleFormatElement appliesTo)
        {
            _appliesTo = appliesTo;
            return this;
        }
    }

    internal class XLPivotFormatList : List<IXLPivotFormat>, IXLPivotFormatList
    {
        public IXLPivotFormat Add(Action<PivotFormatFactoryProvider> config)
        {
            var provider = new PivotFormatFactoryProvider(XLStyle.Default);
            config(provider);
            provider.Builder.Build();
            Add(provider.Format);
            return provider.Format;
        }

        void IXLPivotFormatList.Add(IXLPivotFormat format)
        {
            Add(format);
        }
    }

    internal class FieldRef : IFieldRef
    {
        public int FieldIdx { get; set; }
        public string FieldName { get; set; }
        public int[] Values { get; set; }
        public List<string> FieldNames { get; private set; }
        public Predicate<object> ValueFilter { get; set; }
        public bool DefaultSubtotal { get; set; }
        public bool SumSubtotal { get; set; }
        public bool CountSubtotal { get; set; }
        public bool CountASubtotal { get; set; }
        public bool AverageSubtotal { get; set; }
        public bool MaxSubtotal { get; set; }
        public bool MinSubtotal { get; set; }
        public bool ApplyProductInSubtotal { get; set; }
        public bool ApplyVarianceInSubtotal { get; set; }
        public bool ApplyVariancePInSubtotal { get; set; }
        public bool ApplyStandardDeviationInSubtotal { get; set; }
        public bool ApplyStandardDeviationPInSubtotal { get; set; }

        private FieldRef()
        {
        }

        public static FieldRef Raw(string fieldName, int[] values)
        {
            return new FieldRef()
            {
                FieldName = fieldName,
                Values = values,
            };
        }

        public static FieldRef ValueField(string fieldName)
        {
            return new FieldRef()
            {
                FieldName = XLConstants.PivotTableValuesSentinalLabel,
                FieldNames = new List<string> {fieldName}
            };
        }

        public static FieldRef Refference(string fieldName, int valueIdx)
        {
            return new FieldRef()
            {
                FieldName = fieldName,
                Values = new[] { valueIdx }
            };
        }

        public static FieldRef ForField(string fieldName, Predicate<object> valueFilter = null)
        {
            return new FieldRef()
            {
                FieldName = fieldName,
                Values = new[] { -1 },
                ValueFilter = valueFilter
            };
        }

        public static FieldRef FromOpenXml(PivotAreaReference pref, XLPivotTable pt)
        {
            string fieldName;
            int[] values;
            var fieldIdx = (int)pref.Field.Value;
            if (fieldIdx == -2)
            {
                fieldName = XLConstants.PivotTableValuesSentinalLabel;
                values = new [] { (int)pref.OfType<FieldItem>().First().Val.Value };
            }
            else
            {
                fieldName = pt.SourceRangeFieldsAvailable.ElementAt(fieldIdx);
                //var cacheField = pivotCacheDefinition.CacheFields.ElementAt(fieldIdx) as CacheField;
                values = pref.Any()
                    ? pref.OfType<FieldItem>().Select(item => (int)item.Val.Value).ToArray()
                    : new[] { -1 };
            }

            var result = Raw(fieldName, values);
            result.DefaultSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.DefaultSubtotal, false);
            result.SumSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.SumSubtotal, false);
            result.CountSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.CountSubtotal, false);
            result.CountASubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.CountASubtotal, false);
            result.AverageSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.AverageSubtotal, false);
            result.MaxSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.MaxSubtotal, false);
            result.MinSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.MinSubtotal, false);
            result.ApplyProductInSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.ApplyProductInSubtotal, false);
            result.ApplyVarianceInSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.ApplyVarianceInSubtotal, false);
            result.ApplyVariancePInSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.ApplyVariancePInSubtotal, false);
            result.ApplyStandardDeviationInSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.ApplyStandardDeviationInSubtotal, false);
            result.ApplyStandardDeviationPInSubtotal = OpenXmlHelper.GetBooleanValueAsBool(pref.ApplyStandardDeviationPInSubtotal, false);
            return result;
        }

        public static explicit operator PivotAreaReference(FieldRef value)
        {
            return new PivotAreaReference
            {
                DefaultSubtotal = OpenXmlHelper.GetBooleanValue(value.DefaultSubtotal, false),
                SumSubtotal = OpenXmlHelper.GetBooleanValue(value.SumSubtotal, false),
                CountSubtotal = OpenXmlHelper.GetBooleanValue(value.CountSubtotal, false),
                CountASubtotal = OpenXmlHelper.GetBooleanValue(value.CountASubtotal, false),
                AverageSubtotal = OpenXmlHelper.GetBooleanValue(value.AverageSubtotal, false),
                MaxSubtotal = OpenXmlHelper.GetBooleanValue(value.MaxSubtotal, false),
                MinSubtotal = OpenXmlHelper.GetBooleanValue(value.MinSubtotal, false),
                ApplyProductInSubtotal = OpenXmlHelper.GetBooleanValue(value.ApplyProductInSubtotal, false),
                ApplyVarianceInSubtotal = OpenXmlHelper.GetBooleanValue(value.ApplyVarianceInSubtotal, false),
                ApplyVariancePInSubtotal = OpenXmlHelper.GetBooleanValue(value.ApplyVariancePInSubtotal, false),
                ApplyStandardDeviationInSubtotal = OpenXmlHelper.GetBooleanValue(value.ApplyStandardDeviationInSubtotal, false),
                ApplyStandardDeviationPInSubtotal = OpenXmlHelper.GetBooleanValue(value.ApplyStandardDeviationPInSubtotal, false),
                Field = UInt32Value.FromUInt32((uint) value.FieldIdx)
            };
        }
    }
}

