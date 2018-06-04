using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLPivotFormat : IXLPivotFormat
    {
        public IXLStyle Style { get; set; }
        public bool DataOnly { get; set; } = true;
        public bool LabelOnly { get; set; } = false;
        public bool Outline { get; set; } = true;
        public int FieldIdx { get; set; } = -1;
        public string FieldName { get; set; }
        public PivotTableAxisValues? Axis { get; set; } = null;
        public PivotAreaValues AreaType { get; set; } = PivotAreaValues.Normal;
        public bool GrandRow { get; set; } = false;
        public bool GrandCol { get; set; } = false;
        public bool CollapsedLevelsAreSubtotals { get; set; } = false;
        public IEnumerable<IFieldRef> FieldReferences { get; } = new List<IFieldRef>();

        public XLPivotFormat(IXLStyle style)
        {
            Style = style;
        }

        public void Build(XLPivotTable pt, XLWorkbook.SaveContext context)
        {
            if (Axis.HasValue && !string.IsNullOrEmpty(FieldName))
            {
                if (pt.ColumnLabels.Contains(FieldName))
                    Axis = PivotTableAxisValues.AxisColumn;
                else if (pt.RowLabels.Contains(FieldName))
                    Axis = PivotTableAxisValues.AxisRow;
                else if (pt.ReportFilters.Contains(FieldName))
                    Axis = PivotTableAxisValues.AxisPage;
                else Axis = PivotTableAxisValues.AxisValues;
            }

            if (!string.IsNullOrEmpty(FieldName))
                FieldIdx = FieldIndexOf(pt.Fields, FieldName);

            foreach (var field in FieldReferences.Cast<FieldRef>())
            {
                if (field.FieldName == XLConstants.PivotTableValuesSentinalLabel)
                {
                    field.FieldIdx = -2;
                    field.Value = ValueIndexOf(pt.Values, (string)field.Value);
                }
                else
                {
                    field.FieldIdx = FieldIndexOf(pt.Fields, field.FieldName);
                    var values = context.PivotTables[pt.Guid].Fields[field.FieldName].DistinctValues.ToList();

                    if (field.ValueFilter != null)
                    {
                        foreach (var filter in values.Select((val, idx) => new {val, idx}).Where(v => field.ValueFilter(v.val)))
                        {
                            field.Value = filter.idx;
                        }
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
                throw new ArgumentNullException(nameof(fieldName), $"Invalid field name {fieldName}.");

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
            format.DataOnly = false;
            format.LabelOnly = true;
            format.Outline = false;
            format.AreaType = PivotAreaValues.Button;
            format.FieldName = _field.SourceName;
            format.Axis = PivotTableAxisValues.AxisValues;
            return format;
        }
    }

    public class DataPivotFormatBuilder : PivotFormatBuilder
    {
        private bool _dataOnly = true;
        private bool _labelOnly = false;
        private readonly List<FieldRef> _refFields = new List<FieldRef>();

        internal DataPivotFormatBuilder(IXLPivotFormat format)
            : base(format)
        {
        }

        public DataPivotFormatBuilder ForValueField(IXLPivotValue valueField)
        {
            _refFields.Add(new FieldRef(XLConstants.PivotTableValuesSentinalLabel, valueField.SourceName));
            return this;
        }

        public DataPivotFormatBuilder AndWith(IXLPivotField field, Predicate<object> predicate = null)
        {
            _refFields.Add(new FieldRef(field.SourceName, -1) { ValueFilter = predicate });
            return this;
        }

        public DataPivotFormatBuilder DataOnly()
        {
            _dataOnly = true;
            _labelOnly = false;
            return this;
        }

        public DataPivotFormatBuilder LabelOnly()
        {
            _dataOnly = false;
            _labelOnly = true;
            return this;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat)Format;

            ((List<IFieldRef>)format.FieldReferences).AddRange(_refFields);
            format.DataOnly = _dataOnly;
            format.LabelOnly = _labelOnly;
            return format;
        }
    }

    public class SubtotalPivotFormatBuilder : PivotFormatBuilder
    {
        private readonly IXLPivotField _field;

        internal SubtotalPivotFormatBuilder(IXLPivotField field, IXLPivotFormat format)
            : base(format)
        {
            _field = field;
        }

        internal override XLPivotFormat Build()
        {
            var format = (XLPivotFormat)Format;
            var fieldIdx = _field.SourceName;
            var fieldRef = new FieldRef(fieldIdx, -1) { DefaultSubtotal = true };
            ((List<IFieldRef>)format.FieldReferences).Add(fieldRef);
            return format;
        }
    }

    public class GrandTotalPivotFormatBuilder : PivotFormatBuilder
    {
        private readonly PivotTableAxisValues _axis;
        private bool _dataOnly = false;
        private bool _labelOnly = false;

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
                format.Axis = PivotTableAxisValues.AxisRow;
            }
            else if (_axis == PivotTableAxisValues.AxisColumn)
            {
                format.GrandCol = true;
                format.Axis = PivotTableAxisValues.AxisColumn;
            }

            format.DataOnly = _dataOnly;
            format.LabelOnly = _labelOnly;
            return format;
        }

        public GrandTotalPivotFormatBuilder DataOnly()
        {
            _dataOnly = true;
            _labelOnly = false;
            return this;
        }

        public GrandTotalPivotFormatBuilder LabelOnly()
        {
            _dataOnly = false;
            _labelOnly = true;
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
    }

    internal class FieldRef : IFieldRef
    {
        public int FieldIdx { get; set; }
        public string FieldName { get; set; }
        public object Value { get; set; }
        public Predicate<object> ValueFilter { get; set; }
        public bool DefaultSubtotal { get; set; }

        public FieldRef(string fieldName, object valueIdx)
        {
            FieldName = fieldName;
            Value = valueIdx;
        }
    }
}
