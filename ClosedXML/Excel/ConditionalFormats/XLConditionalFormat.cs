#nullable disable warnings

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormat : IXLDxfContainer, IXLConditionalFormat
    {
        private readonly XLWorksheet _worksheet;
        private readonly XLRanges _ranges;

        private sealed class NoRangeCfComparer : IEqualityComparer<XLConditionalFormat>
        {
            public bool Equals(XLConditionalFormat xx, XLConditionalFormat yy)
            {
                if (ReferenceEquals(xx, yy)) return true;
                if (ReferenceEquals(xx, null)) return false;
                if (ReferenceEquals(yy, null)) return false;
                if (xx.GetType() != yy.GetType()) return false;

                var xxValues = xx.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var yyValues = yy.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var xxFormulas = xx.Ranges.Count > 0 ? xx.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XLCell)xx.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)) : null;
                var yyFormulas = yy.Ranges.Count > 0 ? yy.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XLCell)yy.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)) : null;
                var xStyle = xx.FormatValue;
                var yStyle = yy.FormatValue;
                return Equals(xStyle, yStyle)
                    && xx.ConditionalFormatType == yy.ConditionalFormatType
                    && xx.TimePeriod == yy.TimePeriod
                    && xx.IconSetStyle == yy.IconSetStyle
                    && xx.Operator == yy.Operator
                    && xx.Bottom == yy.Bottom
                    && xx.Percent == yy.Percent
                    && xx.ReverseIconOrder == yy.ReverseIconOrder
                    && xx.StopIfTrue == yy.StopIfTrue
                    && xx.ShowIconOnly == yy.ShowIconOnly
                    && xx.ShowBarOnly == yy.ShowBarOnly
                    && SetEquals(xxValues, yyValues)
                    && SetEquals(xxFormulas, yyFormulas)
                    && Equals(xx.Colors, yy.Colors)
                    && Equals(xx.ContentTypes, yy.ContentTypes)
                    && Equals(xx.IconSetOperators, yy.IconSetOperators);
            }

            public int GetHashCode(XLConditionalFormat obj)
            {
                var xx = obj;
                var xStyle = obj.FormatValue;
                var xValues = xx.Values.Values.Where(v => !v.IsFormula).Select(v => v.Value);
                if (obj.Ranges.Count > 0)
                    xValues = xValues
                    .Union(xx.Values.Values.Where(v => v.IsFormula).Select(f => ((XLCell)obj.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)));

                unchecked
                {
                    var hashCode = xStyle.GetHashCode();
                    hashCode = (hashCode * 397) ^ xValues.GetHashCode();
                    hashCode = (hashCode * 397) ^ (xx.Colors != null ? xx.Colors.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (xx.ContentTypes != null ? xx.ContentTypes.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (xx.IconSetOperators != null ? xx.IconSetOperators.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (int)xx.ConditionalFormatType;
                    hashCode = (hashCode * 397) ^ (int)xx.TimePeriod;
                    hashCode = (hashCode * 397) ^ (int)xx.IconSetStyle;
                    hashCode = (hashCode * 397) ^ (int)xx.Operator;
                    hashCode = (hashCode * 397) ^ xx.Bottom.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.Percent.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ReverseIconOrder.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ShowIconOnly.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ShowBarOnly.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.StopIfTrue.GetHashCode();
                    return hashCode;
                }
            }

            private static bool SetEquals<T>(IEnumerable<T> first, IEnumerable<T> second)
            {
                return new HashSet<T>(second, EqualityComparer<T>.Default)
                    .SetEquals(first);
            }

            private static bool Equals<TValue>(Dictionary<int, TValue> x, Dictionary<int, TValue> y)
            {
                if (x.Count != y.Count)
                    return false;
                if (x.Keys.Except(y.Keys).Any())
                    return false;
                if (y.Keys.Except(x.Keys).Any())
                    return false;
                var valueComparer = EqualityComparer<TValue>.Default;
                foreach (var pair in x)
                    if (!valueComparer.Equals(pair.Value, y[pair.Key]))
                        return false;

                return true;
            }
        }

        private void AdjustFormulas(XLCell baseCell, XLCell targetCell)
        {
            var keys = Values.Keys.ToList();
            foreach (var key in keys)
            {
                if (Values[key] == null || !Values[key].IsFormula)
                    continue;

                var r1c1 = baseCell.GetFormulaR1C1(Values[key].Value);
                Values[key] = new XLFormula { _value = targetCell.GetFormulaA1(r1c1), IsFormula = true };
            }
        }

        internal static IEqualityComparer<XLConditionalFormat> NoRangeComparer { get; } = new NoRangeCfComparer();

        #region Constructors

        private XLConditionalFormat(XLWorksheet worksheet)
#if !STYLES_REWORK
            : base(XLStyle.Default.Value)
#endif
        {
            _worksheet = worksheet;
            Id = Guid.NewGuid();
            _ranges = new XLRanges(worksheet);
            Values = new XLDictionary<XLFormula>();
            Colors = new XLDictionary<XLColor>();
            ContentTypes = new XLDictionary<XLCFContentType>();
            IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
        }

        public XLConditionalFormat(XLWorksheet worksheet, XLRange range)
            : this(worksheet)
        {
            if (range != null)
                Ranges.Add(range);
        }

        public XLConditionalFormat(XLWorksheet worksheet, IEnumerable<XLRange> ranges)
            : this(worksheet)
        {
            ranges?.ForEach(range => Ranges.Add(range));
        }

        internal XLConditionalFormat(XLWorksheet worksheet, XLConditionalFormat conditionalFormat, XLAreaList areaList)
            : this(worksheet)
        {
            areaList.ForEach(range => Ranges.Add(_worksheet.Range(range)));
            CopyFrom(conditionalFormat);

            var sourceAnchor = _worksheet.Cell(conditionalFormat.Areas[0].FirstPoint);
            var targetAnchor = _worksheet.Cell(areaList[0].FirstPoint);
            AdjustFormulas(sourceAnchor, targetAnchor);
        }

        #endregion Constructors

        public Guid Id { get; internal set; }

        /// <summary>
        /// Priority of formatting rule. Lower values have higher priority than higher values.
        /// Minimum value is 1. It is basically used for ordering of CF during saving.
        /// </summary>
        internal Int32 Priority { get; set; }

        public XLDxfValue? FormatValue { get; set; }

        internal XLDxFormat Format => new(_worksheet.Workbook.Styles, this);

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetValue(value);
        }

        public XLDictionary<XLFormula> Values { get; private set; }

        public XLDictionary<XLColor> Colors { get; private set; }

        public XLDictionary<XLCFContentType> ContentTypes { get; private set; }

        public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; private set; }

        public IXLRange Range
        {
            get { return Ranges.FirstOrDefault(); }
            set
            {
                Ranges.RemoveAll();
                Ranges.Add(value);
            }
        }

        public IXLRanges Ranges => _ranges;

        public XLConditionalFormatType ConditionalFormatType { get; set; }

        public XLTimePeriod TimePeriod { get; set; }

        public XLIconSetStyle IconSetStyle { get; set; }

        public XLCFOperator Operator { get; set; }

        public Boolean Bottom { get; set; }

        public Boolean Percent { get; set; }

        public Boolean ReverseIconOrder { get; set; }

        public Boolean ShowIconOnly { get; set; }

        public Boolean ShowBarOnly { get; set; }

        public Boolean StopIfTrue { get; set; }

        internal XLAreaList Areas
        {
            get
            {
                return new XLAreaList(_ranges.Select<XLRange, XLSheetRange>(range => range.SheetRange).ToList());
            }
            set
            {
                Ranges.RemoveAll();
                foreach (var area in value)
                    Ranges.Add(_worksheet.Range(area));
            }
        }

        public IXLConditionalFormat SetStopIfTrue()
        {
            return SetStopIfTrue(true);
        }

        public IXLConditionalFormat SetStopIfTrue(bool value)
        {
            this.StopIfTrue = value;
            return this;
        }

        public IXLConditionalFormat CopyTo(IXLWorksheet targetSheet)
        {
            if (targetSheet == Range?.Worksheet)
                throw new InvalidOperationException("Cannot copy conditional format to the worksheet it already belongs to.");
            var newCf = new XLConditionalFormat((XLWorksheet)targetSheet, this, Areas);
            targetSheet.ConditionalFormats.Add(newCf);
            return newCf;
        }

        public void CopyFrom(IXLConditionalFormat other)
        {
#if STYLES_REWORK
            var otherDxf = ((XLConditionalFormat)other).FormatValue;
            FormatValue = otherDxf is not null
                ? _worksheet.Workbook.Styles.GetRegisteredDxFormat(otherDxf, static x => x)
                : null;
#else
            InnerStyle = other.Style;
#endif
            ConditionalFormatType = other.ConditionalFormatType;
            TimePeriod = other.TimePeriod;
            IconSetStyle = other.IconSetStyle;
            Operator = other.Operator;
            Bottom = other.Bottom;
            Percent = other.Percent;
            ReverseIconOrder = other.ReverseIconOrder;
            ShowIconOnly = other.ShowIconOnly;
            ShowBarOnly = other.ShowBarOnly;
            StopIfTrue = other.StopIfTrue;

            Values.Clear();
            other.Values.ForEach(kp => Values.Add(kp.Key, new XLFormula(kp.Value)));
            //CopyDictionary(Values, other.Values);
            CopyDictionary(Colors, other.Colors);
            CopyDictionary(ContentTypes, other.ContentTypes);
            CopyDictionary(IconSetOperators, other.IconSetOperators);
        }

        private void CopyDictionary<T>(XLDictionary<T> target, XLDictionary<T> source)
        {
            target.Clear();
            source.ForEach(kp => target.Add(kp.Key, kp.Value));
        }

        public IXLStyle WhenIsBlank()
        {
            ConditionalFormatType = XLConditionalFormatType.IsBlank;
            return Style;
        }

        public IXLStyle WhenNotBlank()
        {
            ConditionalFormatType = XLConditionalFormatType.NotBlank;
            return Style;
        }

        public IXLStyle WhenIsError()
        {
            ConditionalFormatType = XLConditionalFormatType.IsError;
            return Style;
        }

        public IXLStyle WhenNotError()
        {
            ConditionalFormatType = XLConditionalFormatType.NotError;
            return Style;
        }

        public IXLStyle WhenDateIs(XLTimePeriod timePeriod)
        {
            TimePeriod = timePeriod;
            ConditionalFormatType = XLConditionalFormatType.TimePeriod;
            return Style;
        }

        public IXLStyle WhenContains(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            ConditionalFormatType = XLConditionalFormatType.ContainsText;
            Operator = XLCFOperator.Contains;
            return Style;
        }

        public IXLStyle WhenNotContains(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            ConditionalFormatType = XLConditionalFormatType.NotContainsText;
            Operator = XLCFOperator.NotContains;
            return Style;
        }

        public IXLStyle WhenStartsWith(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            ConditionalFormatType = XLConditionalFormatType.StartsWith;
            Operator = XLCFOperator.StartsWith;
            return Style;
        }

        public IXLStyle WhenEndsWith(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            ConditionalFormatType = XLConditionalFormatType.EndsWith;
            Operator = XLCFOperator.EndsWith;
            return Style;
        }

        public IXLStyle WhenEquals(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotEquals(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenGreaterThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenLessThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrGreaterThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrLessThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenBetween(String minValue, String maxValue)
        {
            Values.Initialize(new XLFormula { Value = minValue });
            Values.Add(new XLFormula { Value = maxValue });
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotBetween(String minValue, String maxValue)
        {
            Values.Initialize(new XLFormula { Value = minValue });
            Values.Add(new XLFormula { Value = maxValue });
            Operator = XLCFOperator.NotBetween;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEquals(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotEquals(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenGreaterThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenLessThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrGreaterThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrLessThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(new XLFormula(minValue));
            Values.Add(new XLFormula(maxValue));
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(new XLFormula(minValue));
            Values.Add(new XLFormula(maxValue));
            Operator = XLCFOperator.NotBetween;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenIsDuplicate()
        {
            ConditionalFormatType = XLConditionalFormatType.IsDuplicate;
            return Style;
        }

        public IXLStyle WhenIsUnique()
        {
            ConditionalFormatType = XLConditionalFormatType.IsUnique;
            return Style;
        }

        public IXLStyle WhenIsTrue(String formula)
        {
            String f = formula.TrimStart()[0] == '=' ? formula : "=" + formula;
            Values.Initialize(new XLFormula { Value = f });
            ConditionalFormatType = XLConditionalFormatType.Expression;
            return Style;
        }

        public IXLStyle WhenIsTop(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(new XLFormula(value));
            Percent = topBottomType == XLTopBottomType.Percent;
            ConditionalFormatType = XLConditionalFormatType.Top10;
            Bottom = false;
            return Style;
        }

        public IXLStyle WhenIsBottom(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(new XLFormula(value));
            Percent = topBottomType == XLTopBottomType.Percent;
            ConditionalFormatType = XLConditionalFormatType.Top10;
            Bottom = true;
            return Style;
        }

        public IXLCFColorScaleMin ColorScale()
        {
            ConditionalFormatType = XLConditionalFormatType.ColorScale;
            return new XLCFColorScaleMin(this);
        }

        public IXLCFDataBarMin DataBar(XLColor color, Boolean showBarOnly = false)
        {
            Colors.Initialize(color);
            ShowBarOnly = showBarOnly;
            ConditionalFormatType = XLConditionalFormatType.DataBar;
            return new XLCFDataBarMin(this);
        }

        public IXLCFDataBarMin DataBar(XLColor positiveColor, XLColor negativeColor, Boolean showBarOnly = false)
        {
            Colors.Initialize(positiveColor);
            Colors.Add(negativeColor);
            ShowBarOnly = showBarOnly;
            ConditionalFormatType = XLConditionalFormatType.DataBar;
            return new XLCFDataBarMin(this);
        }

        public IXLCFIconSet IconSet(XLIconSetStyle iconSetStyle, Boolean reverseIconOrder = false, Boolean showIconOnly = false)
        {
            IconSetOperators.Clear();
            Values.Clear();
            ContentTypes.Clear();
            ConditionalFormatType = XLConditionalFormatType.IconSet;
            IconSetStyle = iconSetStyle;
            ReverseIconOrder = reverseIconOrder;
            ShowIconOnly = showIconOnly;
            return new XLCFIconSet(this);
        }
    }
}
