using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormat: IXLConditionalFormat, IXLStylized
    {
        
        public XLConditionalFormat(XLRange range)
        {
            Range = range;
            Style = new XLStyle(this, range.Worksheet.Style);
            Values = new XLDictionary<String>();
            Colors = new XLDictionary<XLColor>();
            ContentTypes = new XLDictionary<XLCFContentType>();
            IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
        }
        private IXLStyle _style;
        private Int32 _styleCacheId;
        public IXLStyle Style { get { return GetStyle(); } set { SetStyle(value); } }
        private IXLStyle GetStyle()
        {
            //return _style ?? (_style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId)));
            if (_style != null)
                return _style;

            return _style = new XLStyle(this, Range.Worksheet.Workbook.GetStyleById(_styleCacheId));
        }
        private void SetStyle(IXLStyle styleToUse)
        {
            _styleCacheId = Range.Worksheet.Workbook.GetStyleId(styleToUse);
            _style = null;
            StyleChanged = false;
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return Style;
                UpdatingStyle = false;
            }
        }

        public bool UpdatingStyle { get; set; }

        public IXLStyle InnerStyle { get; set; }

        public IXLRanges RangesUsed
        {
            get { return new XLRanges(); }
        }

        public bool StyleChanged { get; set; }
        public XLDictionary<String> Values { get; private set; }
        public XLDictionary<XLColor> Colors { get; private set; }
        public XLDictionary<XLCFContentType> ContentTypes { get; private set; }
        public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; private set; }

        public IXLRange Range { get; set; }
        public XLConditionalFormatType ConditionalFormatType { get; set; }
        public XLTimePeriod TimePeriod { get; set; }
        public XLIconSetStyle IconSetStyle { get; set; }
        public XLCFOperator Operator { get;  set; }
        public Boolean Bottom { get;  set; }
        public Boolean Percent { get;  set; }
        public Boolean ReverseIconOrder { get;  set; }
        public Boolean ShowIconOnly { get;  set; }
        public Boolean ShowBarOnly { get;  set; }

        public void CopyFrom(IXLConditionalFormat other)
        {
            Style = other.Style;
            ConditionalFormatType = other.ConditionalFormatType;
            TimePeriod = other.TimePeriod;
            IconSetStyle = other.IconSetStyle;
            Operator = other.Operator;
            Bottom = other.Bottom;
            Percent = other.Percent;
            ReverseIconOrder = other.ReverseIconOrder;
            ShowIconOnly = other.ShowIconOnly;
            ShowBarOnly = other.ShowBarOnly;
            
            CopyDictionary(Values, other.Values);
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
            Values.Initialize(value);
            ConditionalFormatType = XLConditionalFormatType.ContainsText;
            Operator = XLCFOperator.Contains;
            return Style;
        }
        public IXLStyle WhenNotContains(String value)
        {
            Values.Initialize(value);
            ConditionalFormatType = XLConditionalFormatType.NotContainsText;
            Operator = XLCFOperator.NotContains;
            return Style;
        }
        public IXLStyle WhenStartsWith(String value)
        {
            Values.Initialize(value);
            ConditionalFormatType = XLConditionalFormatType.StartsWith;
            Operator = XLCFOperator.StartsWith;
            return Style;
        }
        public IXLStyle WhenEndsWith(String value)
        {
            Values.Initialize(value);
            ConditionalFormatType = XLConditionalFormatType.EndsWith;
            Operator = XLCFOperator.EndsWith;
            return Style;
        }

        private String GetProperValue(String value)
        {
            return String.Format("\"{0}\"", value.Replace("\"", "\"\""));
        }

        public IXLStyle WhenEquals(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenNotEquals(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenGreaterThan(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenLessThan(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenEqualOrGreaterThan(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenEqualOrLessThan(String value)
        {
            Values.Initialize(GetProperValue(value));
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenBetween(String minValue, String maxValue)
        {
            Values.Initialize(GetProperValue(minValue));
            Values.Add(GetProperValue(maxValue));
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenNotBetween(String minValue, String maxValue)
        {
            Values.Initialize(GetProperValue(minValue));
            Values.Add(GetProperValue(maxValue));
            Operator = XLCFOperator.NotBetween;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEquals(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenNotEquals(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenGreaterThan(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenLessThan(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenEqualOrGreaterThan(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenEqualOrLessThan(Double value)
        {
            Values.Initialize(value.ToString());
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(minValue.ToString());
            Values.Add(maxValue.ToString());
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }
        public IXLStyle WhenNotBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(minValue.ToString());
            Values.Add(maxValue.ToString());
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
            Values.Initialize(formula);
            ConditionalFormatType = XLConditionalFormatType.Expression;
            return Style;
        }
        public IXLStyle WhenIsTop(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(value.ToString());
            Percent = topBottomType == XLTopBottomType.Percent;
            ConditionalFormatType = XLConditionalFormatType.Top10;
            Bottom = false;
            return Style;
        }
        public IXLStyle WhenIsBottom(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(value.ToString());
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

