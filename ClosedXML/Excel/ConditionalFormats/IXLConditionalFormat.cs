namespace ClosedXML.Excel
{
    public enum XLTimePeriod
    {
        Yesterday,
        Today,
        Tomorrow,
        InTheLast7Days,
        LastWeek,
        ThisWeek,
        NextWeek,
        LastMonth,
        ThisMonth,
        NextMonth
    }

    public enum XLIconSetStyle
    {
        ThreeArrows,
        ThreeArrowsGray,
        ThreeFlags,
        ThreeTrafficLights1,
        ThreeTrafficLights2,
        ThreeSigns,
        ThreeSymbols,
        ThreeSymbols2,
        FourArrows,
        FourArrowsGray,
        FourRedToBlack,
        FourRating,
        FourTrafficLights,
        FiveArrows,
        FiveArrowsGray,
        FiveRating,
        FiveQuarters
    }

    public enum XLConditionalFormatType
    {
        Expression,
        CellIs,
        ColorScale,
        DataBar,
        IconSet,
        Top10,
        IsUnique,
        IsDuplicate,
        ContainsText,
        NotContainsText,
        StartsWith,
        EndsWith,
        IsBlank,
        NotBlank,
        IsError,
        NotError,
        TimePeriod,
        AboveAverage
    }

    public enum XLCFOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan, Between, NotBetween, Contains, NotContains, StartsWith, EndsWith }

    public interface IXLConditionalFormat
    {
        IXLStyle Style { get; set; }

        IXLStyle WhenIsBlank();

        IXLStyle WhenNotBlank();

        IXLStyle WhenIsError();

        IXLStyle WhenNotError();

        IXLStyle WhenDateIs(XLTimePeriod timePeriod);

        IXLStyle WhenContains(string value);

        IXLStyle WhenNotContains(string value);

        IXLStyle WhenStartsWith(string value);

        IXLStyle WhenEndsWith(string value);

        IXLStyle WhenEquals(string value);

        IXLStyle WhenNotEquals(string value);

        IXLStyle WhenGreaterThan(string value);

        IXLStyle WhenLessThan(string value);

        IXLStyle WhenEqualOrGreaterThan(string value);

        IXLStyle WhenEqualOrLessThan(string value);

        IXLStyle WhenBetween(string minValue, string maxValue);

        IXLStyle WhenNotBetween(string minValue, string maxValue);

        IXLStyle WhenEquals(double value);

        IXLStyle WhenNotEquals(double value);

        IXLStyle WhenGreaterThan(double value);

        IXLStyle WhenLessThan(double value);

        IXLStyle WhenEqualOrGreaterThan(double value);

        IXLStyle WhenEqualOrLessThan(double value);

        IXLStyle WhenBetween(double minValue, double maxValue);

        IXLStyle WhenNotBetween(double minValue, double maxValue);

        IXLStyle WhenIsDuplicate();

        IXLStyle WhenIsUnique();

        IXLStyle WhenIsTrue(string formula);

        IXLStyle WhenIsTop(int value, XLTopBottomType topBottomType = XLTopBottomType.Items);

        IXLStyle WhenIsBottom(int value, XLTopBottomType topBottomType);

        IXLCFColorScaleMin ColorScale();

        IXLCFDataBarMin DataBar(XLColor color, bool showBarOnly = false);

        IXLCFDataBarMin DataBar(XLColor positiveColor, XLColor negativeColor, bool showBarOnly = false);

        IXLCFIconSet IconSet(XLIconSetStyle iconSetStyle, bool reverseIconOrder = false, bool showIconOnly = false);

        XLConditionalFormatType ConditionalFormatType { get; }

        XLIconSetStyle IconSetStyle { get; }

        XLTimePeriod TimePeriod { get; }

        bool ReverseIconOrder { get; }

        bool ShowIconOnly { get; }

        bool ShowBarOnly { get; }

        bool StopIfTrue { get; }

        /// <summary>
        /// The first of the <see cref="Ranges"/>.
        /// </summary>
        IXLRange Range { get; set; }

        IXLRanges Ranges { get; }

        XLDictionary<XLFormula> Values { get; }

        XLDictionary<XLColor> Colors { get; }

        XLDictionary<XLCFContentType> ContentTypes { get; }

        XLDictionary<XLCFIconSetOperator> IconSetOperators { get; }

        XLCFOperator Operator { get; }

        bool Bottom { get; }

        bool Percent { get; }

        IXLConditionalFormat SetStopIfTrue();

        IXLConditionalFormat SetStopIfTrue(bool value);

        IXLConditionalFormat CopyTo(IXLWorksheet targetSheet);
    }
}
