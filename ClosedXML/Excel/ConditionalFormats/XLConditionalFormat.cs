using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormat : XLStylizedBase, IXLConditionalFormat, IXLStylized
    {
        private sealed class FullEqualityComparer : IEqualityComparer<IXLConditionalFormat>
        {
            private readonly bool _compareRange;
            private readonly DictionaryComparer<int, XLColor> _colorsComparer = new DictionaryComparer<int, XLColor>();
            private readonly EnumerableComparer<string> _listComparer = new EnumerableComparer<string>();
            private readonly DictionaryComparer<int, XLCFContentType> _contentsTypeComparer = new DictionaryComparer<int, XLCFContentType>();
            private readonly DictionaryComparer<int, XLCFIconSetOperator> _iconSetTypeComparer = new DictionaryComparer<int, XLCFIconSetOperator>();

            public FullEqualityComparer(bool compareRange)
            {
                _compareRange = compareRange;
            }

            public bool Equals(IXLConditionalFormat x, IXLConditionalFormat y)
            {
                var xx = (XLConditionalFormat)x;
                var yy = (XLConditionalFormat)y;
                if (ReferenceEquals(xx, yy)) return true;
                if (ReferenceEquals(xx, null)) return false;
                if (ReferenceEquals(yy, null)) return false;
                if (xx.GetType() != yy.GetType()) return false;

                var xxValues = xx.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var yyValues = yy.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var xxFormulas = x.Ranges.Count > 0 ? xx.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XLCell)x.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)) : null;
                var yyFormulas = y.Ranges.Count > 0 ? yy.Values.Values.Where(v => v != null && v.IsFormula).Select(f => ((XLCell)y.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)) : null;

                var xStyle = xx.StyleValue;
                var yStyle = yy.StyleValue;

                return Equals(xStyle, yStyle)
                    && xx.CopyDefaultModify == yy.CopyDefaultModify
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
                    && _listComparer.Equals(xxValues, yyValues)
                    && _listComparer.Equals(xxFormulas, yyFormulas)
                    && _colorsComparer.Equals(xx.Colors, yy.Colors)
                    && _contentsTypeComparer.Equals(xx.ContentTypes, yy.ContentTypes)
                    && _iconSetTypeComparer.Equals(xx.IconSetOperators, yy.IconSetOperators)
                    && (!_compareRange || XLRanges.Equals(xx.Ranges, yy.Ranges));
            }

            public int GetHashCode(IXLConditionalFormat obj)
            {
                var xx = (XLConditionalFormat)obj;
                var xStyle = (obj.Style as XLStyle).Value;
                var xValues = xx.Values.Values.Where(v => !v.IsFormula).Select(v => v.Value);
                if (obj.Ranges.Count > 0)
                    xValues = xValues
                    .Union(xx.Values.Values.Where(v => v.IsFormula).Select(f => ((XLCell)obj.Ranges.First().FirstCell()).GetFormulaR1C1(f.Value)));

                unchecked
                {
                    var hashCode = xStyle.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.StyleValue.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.CopyDefaultModify.GetHashCode();
                    hashCode = (hashCode * 397) ^ xValues.GetHashCode();
                    hashCode = (hashCode * 397) ^ (xx.Colors != null ? xx.Colors.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (xx.ContentTypes != null ? xx.ContentTypes.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (xx.IconSetOperators != null ? xx.IconSetOperators.GetHashCode() : 0);
                    hashCode = (hashCode * 397) ^ (_compareRange && xx.Ranges != null ? xx.Ranges.GetHashCode() : 0);
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
        }

        internal void AdjustFormulas(XLCell baseCell, XLCell targetCell)
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

        private static readonly IEqualityComparer<IXLConditionalFormat> FullComparerInstance = new FullEqualityComparer(true);

        public static IEqualityComparer<IXLConditionalFormat> FullComparer
        {
            get { return FullComparerInstance; }
        }

        private static readonly IEqualityComparer<IXLConditionalFormat> NoRangeComparerInstance = new FullEqualityComparer(false);

        public static IEqualityComparer<IXLConditionalFormat> NoRangeComparer
        {
            get { return NoRangeComparerInstance; }
        }

        #region Constructors

        private XLConditionalFormat(XLStyleValue style)
            : base(XLStyle.Default.Value)
        {
            Id = Guid.NewGuid();
            Ranges = new XLRanges();
            Values = new XLDictionary<XLFormula>();
            Colors = new XLDictionary<XLColor>();
            ContentTypes = new XLDictionary<XLCFContentType>();
            IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
        }

        public XLConditionalFormat(XLRange range, Boolean copyDefaultModify = false)
            : this(XLStyle.Default.Value)
        {
            if (range != null)
                Ranges.Add(range);
            CopyDefaultModify = copyDefaultModify;
        }

        public XLConditionalFormat(IEnumerable<XLRange> ranges, Boolean copyDefaultModify = false)
            : this(XLStyle.Default.Value)
        {
            ranges?.ForEach(range => Ranges.Add(range));
            CopyDefaultModify = copyDefaultModify;
        }

        public XLConditionalFormat(XLConditionalFormat conditionalFormat, IXLRange targetRange)
            : this(conditionalFormat, new[] { targetRange })
        {
        }

        public XLConditionalFormat(XLConditionalFormat conditionalFormat, IEnumerable<IXLRange> targetRanges)
            : this(conditionalFormat.StyleValue)
        {
            targetRanges?.ForEach(range => Ranges.Add(range));
            CopyFrom(conditionalFormat);
        }

        #endregion Constructors

        public Guid Id { get; internal set; }

        internal Int32 OriginalPriority { get; set; }

        public Boolean CopyDefaultModify { get; set; }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get { yield break; }
        }

        public override IXLRanges RangesUsed
        {
            get { return new XLRanges(); }
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

        public IXLRanges Ranges { get; private set; }

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
            var targetRanges = Ranges.Select(r => targetSheet.Range(((XLRangeAddress)r.RangeAddress).WithoutWorksheet()));
            var newCf = new XLConditionalFormat(this, targetRanges);
            targetSheet.ConditionalFormats.Add(newCf);
            return newCf;
        }

        public void CopyFrom(IXLConditionalFormat other)
        {
            InnerStyle = other.Style;
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
            other.Values.Where(x=>x.Value != null).ForEach(kp => Values.Add(kp.Key, new XLFormula(kp.Value)));
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

    internal class DictionaryComparer<TKey, TValue> :
        IEqualityComparer<Dictionary<TKey, TValue>>
    {
        private readonly IEqualityComparer<TValue> _valueComparer;

        public DictionaryComparer(IEqualityComparer<TValue> valueComparer = null)
        {
            this._valueComparer = valueComparer ?? EqualityComparer<TValue>.Default;
        }

        public bool Equals(Dictionary<TKey, TValue> x, Dictionary<TKey, TValue> y)
        {
            if (x.Count != y.Count)
                return false;
            if (x.Keys.Except(y.Keys).Any())
                return false;
            if (y.Keys.Except(x.Keys).Any())
                return false;
            foreach (var pair in x)
                if (!_valueComparer.Equals(pair.Value, y[pair.Key]))
                    return false;
            return true;
        }

        public int GetHashCode(Dictionary<TKey, TValue> obj)
        {
            throw new NotImplementedException();
        }
    }

    internal class EnumerableComparer<T> : IEqualityComparer<IEnumerable<T>>
    {
        private readonly IEqualityComparer<T> _valueComparer;

        public EnumerableComparer(IEqualityComparer<T> valueComparer = null)
        {
            this._valueComparer = valueComparer ?? EqualityComparer<T>.Default;
        }

        public bool Equals(IEnumerable<T> x, IEnumerable<T> y)
        {
            return SetEquals(x, y, _valueComparer);
        }

        public int GetHashCode(IEnumerable<T> obj)
        {
            throw new NotImplementedException();
        }

        public static bool SetEquals(IEnumerable<T> first, IEnumerable<T> second,
            IEqualityComparer<T> comparer)
        {
            return new HashSet<T>(second, comparer ?? EqualityComparer<T>.Default)
                .SetEquals(first);
        }
    }
}
