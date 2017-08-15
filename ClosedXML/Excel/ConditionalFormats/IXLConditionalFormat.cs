using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
        IXLStyle WhenDateIs(XLTimePeriod timePeriod );
        IXLStyle WhenContains(String value);
        IXLStyle WhenNotContains(String value);
        IXLStyle WhenStartsWith(String value);
        IXLStyle WhenEndsWith(String value);
        IXLStyle WhenEquals(String value);
        IXLStyle WhenNotEquals(String value);
        IXLStyle WhenGreaterThan(String value);
        IXLStyle WhenLessThan(String value);
        IXLStyle WhenEqualOrGreaterThan(String value);
        IXLStyle WhenEqualOrLessThan(String value);
        IXLStyle WhenBetween(String minValue, String maxValue);
        IXLStyle WhenNotBetween(String minValue, String maxValue);

        IXLStyle WhenEquals(Double value);
        IXLStyle WhenNotEquals(Double value);
        IXLStyle WhenGreaterThan(Double value);
        IXLStyle WhenLessThan(Double value);
        IXLStyle WhenEqualOrGreaterThan(Double value);
        IXLStyle WhenEqualOrLessThan(Double value);
        IXLStyle WhenBetween(Double minValue, Double maxValue);
        IXLStyle WhenNotBetween(Double minValue, Double maxValue);

        IXLStyle WhenIsDuplicate();
        IXLStyle WhenIsUnique();
        IXLStyle WhenIsTrue(String formula);
        IXLStyle WhenIsTop(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items);
        IXLStyle WhenIsBottom(Int32 value, XLTopBottomType topBottomType);

        IXLCFColorScaleMin ColorScale();
        IXLCFDataBarMin DataBar(XLColor color, Boolean showBarOnly = false);
        IXLCFIconSet IconSet(XLIconSetStyle iconSetStyle, Boolean reverseIconOrder = false, Boolean showIconOnly = false);

        XLConditionalFormatType ConditionalFormatType { get; }
        XLIconSetStyle IconSetStyle { get; }
        XLTimePeriod TimePeriod { get; }
        Boolean ReverseIconOrder { get; }
        Boolean ShowIconOnly { get; }
        Boolean ShowBarOnly { get; }
        IXLRange Range { get; set; }

        XLDictionary<XLFormula> Values { get; }
        XLDictionary<XLColor> Colors { get; }
        XLDictionary<XLCFContentType> ContentTypes { get; }
        XLDictionary<XLCFIconSetOperator> IconSetOperators { get; }
        
        XLCFOperator Operator { get; }
        Boolean Bottom { get;  }
        Boolean Percent { get; }
        
        
    }
}
    