using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using Vml = DocumentFormat.OpenXml.Vml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Excel
{
    internal static class EnumConverter
    {
        #region To OpenXml

        public static UnderlineValues ToOpenXml(this XLFontUnderlineValues value)
        {
            switch (value)
            {
                case XLFontUnderlineValues.Double:
                    return UnderlineValues.Double;

                case XLFontUnderlineValues.DoubleAccounting:
                    return UnderlineValues.DoubleAccounting;

                case XLFontUnderlineValues.None:
                    return UnderlineValues.None;

                case XLFontUnderlineValues.Single:
                    return UnderlineValues.Single;

                case XLFontUnderlineValues.SingleAccounting:
                    return UnderlineValues.SingleAccounting;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static OrientationValues ToOpenXml(this XLPageOrientation value)
        {
            switch (value)
            {
                case XLPageOrientation.Default:
                    return OrientationValues.Default;

                case XLPageOrientation.Landscape:
                    return OrientationValues.Landscape;

                case XLPageOrientation.Portrait:
                    return OrientationValues.Portrait;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static VerticalAlignmentRunValues ToOpenXml(this XLFontVerticalTextAlignmentValues value)
        {
            switch (value)
            {
                case XLFontVerticalTextAlignmentValues.Baseline:
                    return VerticalAlignmentRunValues.Baseline;

                case XLFontVerticalTextAlignmentValues.Subscript:
                    return VerticalAlignmentRunValues.Subscript;

                case XLFontVerticalTextAlignmentValues.Superscript:
                    return VerticalAlignmentRunValues.Superscript;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PatternValues ToOpenXml(this XLFillPatternValues value)
        {
            switch (value)
            {
                case XLFillPatternValues.DarkDown:
                    return PatternValues.DarkDown;

                case XLFillPatternValues.DarkGray:
                    return PatternValues.DarkGray;

                case XLFillPatternValues.DarkGrid:
                    return PatternValues.DarkGrid;

                case XLFillPatternValues.DarkHorizontal:
                    return PatternValues.DarkHorizontal;

                case XLFillPatternValues.DarkTrellis:
                    return PatternValues.DarkTrellis;

                case XLFillPatternValues.DarkUp:
                    return PatternValues.DarkUp;

                case XLFillPatternValues.DarkVertical:
                    return PatternValues.DarkVertical;

                case XLFillPatternValues.Gray0625:
                    return PatternValues.Gray0625;

                case XLFillPatternValues.Gray125:
                    return PatternValues.Gray125;

                case XLFillPatternValues.LightDown:
                    return PatternValues.LightDown;

                case XLFillPatternValues.LightGray:
                    return PatternValues.LightGray;

                case XLFillPatternValues.LightGrid:
                    return PatternValues.LightGrid;

                case XLFillPatternValues.LightHorizontal:
                    return PatternValues.LightHorizontal;

                case XLFillPatternValues.LightTrellis:
                    return PatternValues.LightTrellis;

                case XLFillPatternValues.LightUp:
                    return PatternValues.LightUp;

                case XLFillPatternValues.LightVertical:
                    return PatternValues.LightVertical;

                case XLFillPatternValues.MediumGray:
                    return PatternValues.MediumGray;

                case XLFillPatternValues.None:
                    return PatternValues.None;

                case XLFillPatternValues.Solid:
                    return PatternValues.Solid;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static BorderStyleValues ToOpenXml(this XLBorderStyleValues value)
        {
            switch (value)
            {
                case XLBorderStyleValues.DashDot:
                    return BorderStyleValues.DashDot;

                case XLBorderStyleValues.DashDotDot:
                    return BorderStyleValues.DashDotDot;

                case XLBorderStyleValues.Dashed:
                    return BorderStyleValues.Dashed;

                case XLBorderStyleValues.Dotted:
                    return BorderStyleValues.Dotted;

                case XLBorderStyleValues.Double:
                    return BorderStyleValues.Double;

                case XLBorderStyleValues.Hair:
                    return BorderStyleValues.Hair;

                case XLBorderStyleValues.Medium:
                    return BorderStyleValues.Medium;

                case XLBorderStyleValues.MediumDashDot:
                    return BorderStyleValues.MediumDashDot;

                case XLBorderStyleValues.MediumDashDotDot:
                    return BorderStyleValues.MediumDashDotDot;

                case XLBorderStyleValues.MediumDashed:
                    return BorderStyleValues.MediumDashed;

                case XLBorderStyleValues.None:
                    return BorderStyleValues.None;

                case XLBorderStyleValues.SlantDashDot:
                    return BorderStyleValues.SlantDashDot;

                case XLBorderStyleValues.Thick:
                    return BorderStyleValues.Thick;

                case XLBorderStyleValues.Thin:
                    return BorderStyleValues.Thin;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static HorizontalAlignmentValues ToOpenXml(this XLAlignmentHorizontalValues value)
        {
            switch (value)
            {
                case XLAlignmentHorizontalValues.Center:
                    return HorizontalAlignmentValues.Center;

                case XLAlignmentHorizontalValues.CenterContinuous:
                    return HorizontalAlignmentValues.CenterContinuous;

                case XLAlignmentHorizontalValues.Distributed:
                    return HorizontalAlignmentValues.Distributed;

                case XLAlignmentHorizontalValues.Fill:
                    return HorizontalAlignmentValues.Fill;

                case XLAlignmentHorizontalValues.General:
                    return HorizontalAlignmentValues.General;

                case XLAlignmentHorizontalValues.Justify:
                    return HorizontalAlignmentValues.Justify;

                case XLAlignmentHorizontalValues.Left:
                    return HorizontalAlignmentValues.Left;

                case XLAlignmentHorizontalValues.Right:
                    return HorizontalAlignmentValues.Right;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static VerticalAlignmentValues ToOpenXml(this XLAlignmentVerticalValues value)
        {
            switch (value)
            {
                case XLAlignmentVerticalValues.Bottom:
                    return VerticalAlignmentValues.Bottom;

                case XLAlignmentVerticalValues.Center:
                    return VerticalAlignmentValues.Center;

                case XLAlignmentVerticalValues.Distributed:
                    return VerticalAlignmentValues.Distributed;

                case XLAlignmentVerticalValues.Justify:
                    return VerticalAlignmentValues.Justify;

                case XLAlignmentVerticalValues.Top:
                    return VerticalAlignmentValues.Top;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PageOrderValues ToOpenXml(this XLPageOrderValues value)
        {
            switch (value)
            {
                case XLPageOrderValues.DownThenOver:
                    return PageOrderValues.DownThenOver;

                case XLPageOrderValues.OverThenDown:
                    return PageOrderValues.OverThenDown;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static CellCommentsValues ToOpenXml(this XLShowCommentsValues value)
        {
            switch (value)
            {
                case XLShowCommentsValues.AsDisplayed:
                    return CellCommentsValues.AsDisplayed;

                case XLShowCommentsValues.AtEnd:
                    return CellCommentsValues.AtEnd;

                case XLShowCommentsValues.None:
                    return CellCommentsValues.None;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PrintErrorValues ToOpenXml(this XLPrintErrorValues value)
        {
            switch (value)
            {
                case XLPrintErrorValues.Blank:
                    return PrintErrorValues.Blank;

                case XLPrintErrorValues.Dash:
                    return PrintErrorValues.Dash;

                case XLPrintErrorValues.Displayed:
                    return PrintErrorValues.Displayed;

                case XLPrintErrorValues.NA:
                    return PrintErrorValues.NA;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static CalculateModeValues ToOpenXml(this XLCalculateMode value)
        {
            switch (value)
            {
                case XLCalculateMode.Auto:
                    return CalculateModeValues.Auto;

                case XLCalculateMode.AutoNoTable:
                    return CalculateModeValues.AutoNoTable;

                case XLCalculateMode.Manual:
                    return CalculateModeValues.Manual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ReferenceModeValues ToOpenXml(this XLReferenceStyle value)
        {
            switch (value)
            {
                case XLReferenceStyle.R1C1:
                    return ReferenceModeValues.R1C1;

                case XLReferenceStyle.A1:
                    return ReferenceModeValues.A1;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static uint ToOpenXml(this XLAlignmentReadingOrderValues value)
        {
            switch (value)
            {
                case XLAlignmentReadingOrderValues.ContextDependent:
                    return 0;

                case XLAlignmentReadingOrderValues.LeftToRight:
                    return 1;

                case XLAlignmentReadingOrderValues.RightToLeft:
                    return 2;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static TotalsRowFunctionValues ToOpenXml(this XLTotalsRowFunction value)
        {
            switch (value)
            {
                case XLTotalsRowFunction.None:
                    return TotalsRowFunctionValues.None;

                case XLTotalsRowFunction.Sum:
                    return TotalsRowFunctionValues.Sum;

                case XLTotalsRowFunction.Minimum:
                    return TotalsRowFunctionValues.Minimum;

                case XLTotalsRowFunction.Maximum:
                    return TotalsRowFunctionValues.Maximum;

                case XLTotalsRowFunction.Average:
                    return TotalsRowFunctionValues.Average;

                case XLTotalsRowFunction.Count:
                    return TotalsRowFunctionValues.Count;

                case XLTotalsRowFunction.CountNumbers:
                    return TotalsRowFunctionValues.CountNumbers;

                case XLTotalsRowFunction.StandardDeviation:
                    return TotalsRowFunctionValues.StandardDeviation;

                case XLTotalsRowFunction.Variance:
                    return TotalsRowFunctionValues.Variance;

                case XLTotalsRowFunction.Custom:
                    return TotalsRowFunctionValues.Custom;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationValues ToOpenXml(this XLAllowedValues value)
        {
            switch (value)
            {
                case XLAllowedValues.AnyValue:
                    return DataValidationValues.None;

                case XLAllowedValues.Custom:
                    return DataValidationValues.Custom;

                case XLAllowedValues.Date:
                    return DataValidationValues.Date;

                case XLAllowedValues.Decimal:
                    return DataValidationValues.Decimal;

                case XLAllowedValues.List:
                    return DataValidationValues.List;

                case XLAllowedValues.TextLength:
                    return DataValidationValues.TextLength;

                case XLAllowedValues.Time:
                    return DataValidationValues.Time;

                case XLAllowedValues.WholeNumber:
                    return DataValidationValues.Whole;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationErrorStyleValues ToOpenXml(this XLErrorStyle value)
        {
            switch (value)
            {
                case XLErrorStyle.Information:
                    return DataValidationErrorStyleValues.Information;

                case XLErrorStyle.Warning:
                    return DataValidationErrorStyleValues.Warning;

                case XLErrorStyle.Stop:
                    return DataValidationErrorStyleValues.Stop;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationOperatorValues ToOpenXml(this XLOperator value)
        {
            switch (value)
            {
                case XLOperator.Between:
                    return DataValidationOperatorValues.Between;

                case XLOperator.EqualOrGreaterThan:
                    return DataValidationOperatorValues.GreaterThanOrEqual;

                case XLOperator.EqualOrLessThan:
                    return DataValidationOperatorValues.LessThanOrEqual;

                case XLOperator.EqualTo:
                    return DataValidationOperatorValues.Equal;

                case XLOperator.GreaterThan:
                    return DataValidationOperatorValues.GreaterThan;

                case XLOperator.LessThan:
                    return DataValidationOperatorValues.LessThan;

                case XLOperator.NotBetween:
                    return DataValidationOperatorValues.NotBetween;

                case XLOperator.NotEqualTo:
                    return DataValidationOperatorValues.NotEqual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static SheetStateValues ToOpenXml(this XLWorksheetVisibility value)
        {
            switch (value)
            {
                case XLWorksheetVisibility.Visible:
                    return SheetStateValues.Visible;

                case XLWorksheetVisibility.Hidden:
                    return SheetStateValues.Hidden;

                case XLWorksheetVisibility.VeryHidden:
                    return SheetStateValues.VeryHidden;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PhoneticAlignmentValues ToOpenXml(this XLPhoneticAlignment value)
        {
            switch (value)
            {
                case XLPhoneticAlignment.Center:
                    return PhoneticAlignmentValues.Center;

                case XLPhoneticAlignment.Distributed:
                    return PhoneticAlignmentValues.Distributed;

                case XLPhoneticAlignment.Left:
                    return PhoneticAlignmentValues.Left;

                case XLPhoneticAlignment.NoControl:
                    return PhoneticAlignmentValues.NoControl;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PhoneticValues ToOpenXml(this XLPhoneticType value)
        {
            switch (value)
            {
                case XLPhoneticType.FullWidthKatakana:
                    return PhoneticValues.FullWidthKatakana;

                case XLPhoneticType.HalfWidthKatakana:
                    return PhoneticValues.HalfWidthKatakana;

                case XLPhoneticType.Hiragana:
                    return PhoneticValues.Hiragana;

                case XLPhoneticType.NoConversion:
                    return PhoneticValues.NoConversion;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataConsolidateFunctionValues ToOpenXml(this XLPivotSummary value)
        {
            switch (value)
            {
                case XLPivotSummary.Sum: return DataConsolidateFunctionValues.Sum;
                case XLPivotSummary.Count: return DataConsolidateFunctionValues.Count;
                case XLPivotSummary.Average: return DataConsolidateFunctionValues.Average;
                case XLPivotSummary.Minimum: return DataConsolidateFunctionValues.Minimum;
                case XLPivotSummary.Maximum: return DataConsolidateFunctionValues.Maximum;
                case XLPivotSummary.Product: return DataConsolidateFunctionValues.Product;
                case XLPivotSummary.CountNumbers: return DataConsolidateFunctionValues.CountNumbers;
                case XLPivotSummary.StandardDeviation: return DataConsolidateFunctionValues.StandardDeviation;
                case XLPivotSummary.PopulationStandardDeviation: return DataConsolidateFunctionValues.StandardDeviationP;
                case XLPivotSummary.Variance: return DataConsolidateFunctionValues.Variance;
                case XLPivotSummary.PopulationVariance: return DataConsolidateFunctionValues.VarianceP;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ShowDataAsValues ToOpenXml(this XLPivotCalculation value)
        {
            switch (value)
            {
                case XLPivotCalculation.Normal: return ShowDataAsValues.Normal;
                case XLPivotCalculation.DifferenceFrom: return ShowDataAsValues.Difference;
                case XLPivotCalculation.PercentageOf: return ShowDataAsValues.Percent;
                case XLPivotCalculation.PercentageDifferenceFrom: return ShowDataAsValues.PercentageDifference;
                case XLPivotCalculation.RunningTotal: return ShowDataAsValues.RunTotal;
                case XLPivotCalculation.PercentageOfRow: return ShowDataAsValues.PercentOfRaw; // There's a typo in the OpenXML SDK =)
                case XLPivotCalculation.PercentageOfColumn: return ShowDataAsValues.PercentOfColumn;
                case XLPivotCalculation.PercentageOfTotal: return ShowDataAsValues.PercentOfTotal;
                case XLPivotCalculation.Index: return ShowDataAsValues.Index;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static FilterOperatorValues ToOpenXml(this XLFilterOperator value)
        {
            switch (value)
            {
                case XLFilterOperator.Equal: return FilterOperatorValues.Equal;
                case XLFilterOperator.NotEqual: return FilterOperatorValues.NotEqual;
                case XLFilterOperator.GreaterThan: return FilterOperatorValues.GreaterThan;
                case XLFilterOperator.EqualOrGreaterThan: return FilterOperatorValues.GreaterThanOrEqual;
                case XLFilterOperator.LessThan: return FilterOperatorValues.LessThan;
                case XLFilterOperator.EqualOrLessThan: return FilterOperatorValues.LessThanOrEqual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DynamicFilterValues ToOpenXml(this XLFilterDynamicType value)
        {
            switch (value)
            {
                case XLFilterDynamicType.AboveAverage: return DynamicFilterValues.AboveAverage;
                case XLFilterDynamicType.BelowAverage: return DynamicFilterValues.BelowAverage;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DateTimeGroupingValues ToOpenXml(this XLDateTimeGrouping value)
        {
            switch (value)
            {
                case XLDateTimeGrouping.Year: return DateTimeGroupingValues.Year;
                case XLDateTimeGrouping.Month: return DateTimeGroupingValues.Month;
                case XLDateTimeGrouping.Day: return DateTimeGroupingValues.Day;
                case XLDateTimeGrouping.Hour: return DateTimeGroupingValues.Hour;
                case XLDateTimeGrouping.Minute: return DateTimeGroupingValues.Minute;
                case XLDateTimeGrouping.Second: return DateTimeGroupingValues.Second;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static SheetViewValues ToOpenXml(this XLSheetViewOptions value)
        {
            switch (value)
            {
                case XLSheetViewOptions.Normal: return SheetViewValues.Normal;
                case XLSheetViewOptions.PageBreakPreview: return SheetViewValues.PageBreakPreview;
                case XLSheetViewOptions.PageLayout: return SheetViewValues.PageLayout;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static Vml.StrokeLineStyleValues ToOpenXml(this XLLineStyle value)
        {
            switch (value)
            {
                case XLLineStyle.Single: return Vml.StrokeLineStyleValues.Single;
                case XLLineStyle.ThickBetweenThin: return Vml.StrokeLineStyleValues.ThickBetweenThin;
                case XLLineStyle.ThickThin: return Vml.StrokeLineStyleValues.ThickThin;
                case XLLineStyle.ThinThick: return Vml.StrokeLineStyleValues.ThinThick;
                case XLLineStyle.ThinThin: return Vml.StrokeLineStyleValues.ThinThin;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormatValues ToOpenXml(this XLConditionalFormatType value)
        {
            switch (value)
            {
                case XLConditionalFormatType.Expression: return ConditionalFormatValues.Expression;
                case XLConditionalFormatType.CellIs: return ConditionalFormatValues.CellIs;
                case XLConditionalFormatType.ColorScale: return ConditionalFormatValues.ColorScale;
                case XLConditionalFormatType.DataBar: return ConditionalFormatValues.DataBar;
                case XLConditionalFormatType.IconSet: return ConditionalFormatValues.IconSet;
                case XLConditionalFormatType.Top10: return ConditionalFormatValues.Top10;
                case XLConditionalFormatType.IsUnique: return ConditionalFormatValues.UniqueValues;
                case XLConditionalFormatType.IsDuplicate: return ConditionalFormatValues.DuplicateValues;
                case XLConditionalFormatType.ContainsText: return ConditionalFormatValues.ContainsText;
                case XLConditionalFormatType.NotContainsText: return ConditionalFormatValues.NotContainsText;
                case XLConditionalFormatType.StartsWith: return ConditionalFormatValues.BeginsWith;
                case XLConditionalFormatType.EndsWith: return ConditionalFormatValues.EndsWith;
                case XLConditionalFormatType.IsBlank: return ConditionalFormatValues.ContainsBlanks;
                case XLConditionalFormatType.NotBlank: return ConditionalFormatValues.NotContainsBlanks;
                case XLConditionalFormatType.IsError: return ConditionalFormatValues.ContainsErrors;
                case XLConditionalFormatType.NotError: return ConditionalFormatValues.NotContainsErrors;
                case XLConditionalFormatType.TimePeriod: return ConditionalFormatValues.TimePeriod;
                case XLConditionalFormatType.AboveAverage: return ConditionalFormatValues.AboveAverage;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormatValueObjectValues ToOpenXml(this XLCFContentType value)
        {
            switch (value)
            {
                case XLCFContentType.Number: return ConditionalFormatValueObjectValues.Number;
                case XLCFContentType.Percent: return ConditionalFormatValueObjectValues.Percent;
                case XLCFContentType.Maximum: return ConditionalFormatValueObjectValues.Max;
                case XLCFContentType.Minimum: return ConditionalFormatValueObjectValues.Min;
                case XLCFContentType.Formula: return ConditionalFormatValueObjectValues.Formula;
                case XLCFContentType.Percentile: return ConditionalFormatValueObjectValues.Percentile;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormattingOperatorValues ToOpenXml(this XLCFOperator value)
        {
            switch (value)
            {
                case XLCFOperator.LessThan: return ConditionalFormattingOperatorValues.LessThan;
                case XLCFOperator.EqualOrLessThan: return ConditionalFormattingOperatorValues.LessThanOrEqual;
                case XLCFOperator.Equal: return ConditionalFormattingOperatorValues.Equal;
                case XLCFOperator.NotEqual: return ConditionalFormattingOperatorValues.NotEqual;
                case XLCFOperator.EqualOrGreaterThan: return ConditionalFormattingOperatorValues.GreaterThanOrEqual;
                case XLCFOperator.GreaterThan: return ConditionalFormattingOperatorValues.GreaterThan;
                case XLCFOperator.Between: return ConditionalFormattingOperatorValues.Between;
                case XLCFOperator.NotBetween: return ConditionalFormattingOperatorValues.NotBetween;
                case XLCFOperator.Contains: return ConditionalFormattingOperatorValues.ContainsText;
                case XLCFOperator.NotContains: return ConditionalFormattingOperatorValues.NotContains;
                case XLCFOperator.StartsWith: return ConditionalFormattingOperatorValues.BeginsWith;
                case XLCFOperator.EndsWith: return ConditionalFormattingOperatorValues.EndsWith;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static IconSetValues ToOpenXml(this XLIconSetStyle value)
        {
            switch (value)
            {
                case XLIconSetStyle.ThreeArrows: return IconSetValues.ThreeArrows;
                case XLIconSetStyle.ThreeArrowsGray: return IconSetValues.ThreeArrowsGray;
                case XLIconSetStyle.ThreeFlags: return IconSetValues.ThreeFlags;
                case XLIconSetStyle.ThreeTrafficLights1: return IconSetValues.ThreeTrafficLights1;
                case XLIconSetStyle.ThreeTrafficLights2: return IconSetValues.ThreeTrafficLights2;
                case XLIconSetStyle.ThreeSigns: return IconSetValues.ThreeSigns;
                case XLIconSetStyle.ThreeSymbols: return IconSetValues.ThreeSymbols;
                case XLIconSetStyle.ThreeSymbols2: return IconSetValues.ThreeSymbols2;
                case XLIconSetStyle.FourArrows: return IconSetValues.FourArrows;
                case XLIconSetStyle.FourArrowsGray: return IconSetValues.FourArrowsGray;
                case XLIconSetStyle.FourRedToBlack: return IconSetValues.FourRedToBlack;
                case XLIconSetStyle.FourRating: return IconSetValues.FourRating;
                case XLIconSetStyle.FourTrafficLights: return IconSetValues.FourTrafficLights;
                case XLIconSetStyle.FiveArrows: return IconSetValues.FiveArrows;
                case XLIconSetStyle.FiveArrowsGray: return IconSetValues.FiveArrowsGray;
                case XLIconSetStyle.FiveRating: return IconSetValues.FiveRating;
                case XLIconSetStyle.FiveQuarters: return IconSetValues.FiveQuarters;

                default:
                    throw new ArgumentOutOfRangeException("Not implemented value!");
            }
        }

        public static TimePeriodValues ToOpenXml(this XLTimePeriod value)
        {
            switch (value)
            {
                case XLTimePeriod.Yesterday: return TimePeriodValues.Yesterday;
                case XLTimePeriod.Today: return TimePeriodValues.Today;
                case XLTimePeriod.Tomorrow: return TimePeriodValues.Tomorrow;
                case XLTimePeriod.InTheLast7Days: return TimePeriodValues.Last7Days;
                case XLTimePeriod.LastWeek: return TimePeriodValues.LastWeek;
                case XLTimePeriod.ThisWeek: return TimePeriodValues.ThisWeek;
                case XLTimePeriod.NextWeek: return TimePeriodValues.NextWeek;
                case XLTimePeriod.LastMonth: return TimePeriodValues.LastMonth;
                case XLTimePeriod.ThisMonth: return TimePeriodValues.ThisMonth;
                case XLTimePeriod.NextMonth: return TimePeriodValues.NextMonth;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ImagePartType ToOpenXml(this XLPictureFormat value)
        {
            return Enum.Parse(typeof(ImagePartType), value.ToString()).CastTo<ImagePartType>();
        }

        public static Xdr.EditAsValues ToOpenXml(this XLPicturePlacement value)
        {
            switch (value)
            {
                case XLPicturePlacement.FreeFloating:
                    return Xdr.EditAsValues.Absolute;

                case XLPicturePlacement.Move:
                    return Xdr.EditAsValues.OneCell;

                case XLPicturePlacement.MoveAndSize:
                    return Xdr.EditAsValues.TwoCell;

                default:
                    throw new ArgumentOutOfRangeException("Not implemented value!");
            }
        }

        public static PivotAreaValues ToOpenXml(this XLPivotAreaValues value)
        {
            switch (value)
            {
                case XLPivotAreaValues.None:
                    return PivotAreaValues.None;

                case XLPivotAreaValues.Normal:
                    return PivotAreaValues.Normal;

                case XLPivotAreaValues.Data:
                    return PivotAreaValues.Data;

                case XLPivotAreaValues.All:
                    return PivotAreaValues.All;

                case XLPivotAreaValues.Origin:
                    return PivotAreaValues.Origin;

                case XLPivotAreaValues.Button:
                    return PivotAreaValues.Button;

                case XLPivotAreaValues.TopRight:
                    return PivotAreaValues.TopRight;

                case XLPivotAreaValues.TopEnd:
                    return PivotAreaValues.TopEnd;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "XLPivotAreaValues value not implemented");
            }
        }

        public static X14.SparklineTypeValues ToOpenXml(this XLSparklineType value)
        {
            switch (value)
            {
                case XLSparklineType.Line: return X14.SparklineTypeValues.Line;
                case XLSparklineType.Column: return X14.SparklineTypeValues.Column;
                case XLSparklineType.Stacked: return X14.SparklineTypeValues.Stacked;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static X14.SparklineAxisMinMaxValues ToOpenXml(this XLSparklineAxisMinMax value)
        {
            switch (value)
            {
                case XLSparklineAxisMinMax.Automatic: return X14.SparklineAxisMinMaxValues.Individual;
                case XLSparklineAxisMinMax.SameForAll: return X14.SparklineAxisMinMaxValues.Group;
                case XLSparklineAxisMinMax.Custom: return X14.SparklineAxisMinMaxValues.Custom;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static X14.DisplayBlanksAsValues ToOpenXml(this XLDisplayBlanksAsValues value)
        {
            switch (value)
            {
                case XLDisplayBlanksAsValues.Interpolate: return X14.DisplayBlanksAsValues.Span;
                case XLDisplayBlanksAsValues.NotPlotted: return X14.DisplayBlanksAsValues.Gap;
                case XLDisplayBlanksAsValues.Zero: return X14.DisplayBlanksAsValues.Zero;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        #endregion To OpenXml

        #region To ClosedXml

        public static XLFontUnderlineValues ToClosedXml(this UnderlineValues value)
        {
            switch (value)
            {
                case UnderlineValues.Double:
                    return XLFontUnderlineValues.Double;

                case UnderlineValues.DoubleAccounting:
                    return XLFontUnderlineValues.DoubleAccounting;

                case UnderlineValues.None:
                    return XLFontUnderlineValues.None;

                case UnderlineValues.Single:
                    return XLFontUnderlineValues.Single;

                case UnderlineValues.SingleAccounting:
                    return XLFontUnderlineValues.SingleAccounting;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPageOrientation ToClosedXml(this OrientationValues value)
        {
            switch (value)
            {
                case OrientationValues.Default:
                    return XLPageOrientation.Default;

                case OrientationValues.Landscape:
                    return XLPageOrientation.Landscape;

                case OrientationValues.Portrait:
                    return XLPageOrientation.Portrait;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLFontVerticalTextAlignmentValues ToClosedXml(this VerticalAlignmentRunValues value)
        {
            switch (value)
            {
                case VerticalAlignmentRunValues.Baseline:
                    return XLFontVerticalTextAlignmentValues.Baseline;

                case VerticalAlignmentRunValues.Subscript:
                    return XLFontVerticalTextAlignmentValues.Subscript;

                case VerticalAlignmentRunValues.Superscript:
                    return XLFontVerticalTextAlignmentValues.Superscript;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLFillPatternValues ToClosedXml(this PatternValues value)
        {
            switch (value)
            {
                case PatternValues.DarkDown:
                    return XLFillPatternValues.DarkDown;

                case PatternValues.DarkGray:
                    return XLFillPatternValues.DarkGray;

                case PatternValues.DarkGrid:
                    return XLFillPatternValues.DarkGrid;

                case PatternValues.DarkHorizontal:
                    return XLFillPatternValues.DarkHorizontal;

                case PatternValues.DarkTrellis:
                    return XLFillPatternValues.DarkTrellis;

                case PatternValues.DarkUp:
                    return XLFillPatternValues.DarkUp;

                case PatternValues.DarkVertical:
                    return XLFillPatternValues.DarkVertical;

                case PatternValues.Gray0625:
                    return XLFillPatternValues.Gray0625;

                case PatternValues.Gray125:
                    return XLFillPatternValues.Gray125;

                case PatternValues.LightDown:
                    return XLFillPatternValues.LightDown;

                case PatternValues.LightGray:
                    return XLFillPatternValues.LightGray;

                case PatternValues.LightGrid:
                    return XLFillPatternValues.LightGrid;

                case PatternValues.LightHorizontal:
                    return XLFillPatternValues.LightHorizontal;

                case PatternValues.LightTrellis:
                    return XLFillPatternValues.LightTrellis;

                case PatternValues.LightUp:
                    return XLFillPatternValues.LightUp;

                case PatternValues.LightVertical:
                    return XLFillPatternValues.LightVertical;

                case PatternValues.MediumGray:
                    return XLFillPatternValues.MediumGray;

                case PatternValues.None:
                    return XLFillPatternValues.None;

                case PatternValues.Solid:
                    return XLFillPatternValues.Solid;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLBorderStyleValues ToClosedXml(this BorderStyleValues value)
        {
            switch (value)
            {
                case BorderStyleValues.DashDot:
                    return XLBorderStyleValues.DashDot;

                case BorderStyleValues.DashDotDot:
                    return XLBorderStyleValues.DashDotDot;

                case BorderStyleValues.Dashed:
                    return XLBorderStyleValues.Dashed;

                case BorderStyleValues.Dotted:
                    return XLBorderStyleValues.Dotted;

                case BorderStyleValues.Double:
                    return XLBorderStyleValues.Double;

                case BorderStyleValues.Hair:
                    return XLBorderStyleValues.Hair;

                case BorderStyleValues.Medium:
                    return XLBorderStyleValues.Medium;

                case BorderStyleValues.MediumDashDot:
                    return XLBorderStyleValues.MediumDashDot;

                case BorderStyleValues.MediumDashDotDot:
                    return XLBorderStyleValues.MediumDashDotDot;

                case BorderStyleValues.MediumDashed:
                    return XLBorderStyleValues.MediumDashed;

                case BorderStyleValues.None:
                    return XLBorderStyleValues.None;

                case BorderStyleValues.SlantDashDot:
                    return XLBorderStyleValues.SlantDashDot;

                case BorderStyleValues.Thick:
                    return XLBorderStyleValues.Thick;

                case BorderStyleValues.Thin:
                    return XLBorderStyleValues.Thin;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLAlignmentHorizontalValues ToClosedXml(this HorizontalAlignmentValues value)
        {
            switch (value)
            {
                case HorizontalAlignmentValues.Center:
                    return XLAlignmentHorizontalValues.Center;

                case HorizontalAlignmentValues.CenterContinuous:
                    return XLAlignmentHorizontalValues.CenterContinuous;

                case HorizontalAlignmentValues.Distributed:
                    return XLAlignmentHorizontalValues.Distributed;

                case HorizontalAlignmentValues.Fill:
                    return XLAlignmentHorizontalValues.Fill;

                case HorizontalAlignmentValues.General:
                    return XLAlignmentHorizontalValues.General;

                case HorizontalAlignmentValues.Justify:
                    return XLAlignmentHorizontalValues.Justify;

                case HorizontalAlignmentValues.Left:
                    return XLAlignmentHorizontalValues.Left;

                case HorizontalAlignmentValues.Right:
                    return XLAlignmentHorizontalValues.Right;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLAlignmentVerticalValues ToClosedXml(this VerticalAlignmentValues value)
        {
            switch (value)
            {
                case VerticalAlignmentValues.Bottom:
                    return XLAlignmentVerticalValues.Bottom;

                case VerticalAlignmentValues.Center:
                    return XLAlignmentVerticalValues.Center;

                case VerticalAlignmentValues.Distributed:
                    return XLAlignmentVerticalValues.Distributed;

                case VerticalAlignmentValues.Justify:
                    return XLAlignmentVerticalValues.Justify;

                case VerticalAlignmentValues.Top:
                    return XLAlignmentVerticalValues.Top;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPageOrderValues ToClosedXml(this PageOrderValues value)
        {
            switch (value)
            {
                case PageOrderValues.DownThenOver:
                    return XLPageOrderValues.DownThenOver;

                case PageOrderValues.OverThenDown:
                    return XLPageOrderValues.OverThenDown;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLShowCommentsValues ToClosedXml(this CellCommentsValues value)
        {
            switch (value)
            {
                case CellCommentsValues.AsDisplayed:
                    return XLShowCommentsValues.AsDisplayed;

                case CellCommentsValues.AtEnd:
                    return XLShowCommentsValues.AtEnd;

                case CellCommentsValues.None:
                    return XLShowCommentsValues.None;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPrintErrorValues ToClosedXml(this PrintErrorValues value)
        {
            switch (value)
            {
                case PrintErrorValues.Blank:
                    return XLPrintErrorValues.Blank;

                case PrintErrorValues.Dash:
                    return XLPrintErrorValues.Dash;

                case PrintErrorValues.Displayed:
                    return XLPrintErrorValues.Displayed;

                case PrintErrorValues.NA:
                    return XLPrintErrorValues.NA;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLCalculateMode ToClosedXml(this CalculateModeValues value)
        {
            switch (value)
            {
                case CalculateModeValues.Auto:
                    return XLCalculateMode.Auto;

                case CalculateModeValues.AutoNoTable:
                    return XLCalculateMode.AutoNoTable;

                case CalculateModeValues.Manual:
                    return XLCalculateMode.Manual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLReferenceStyle ToClosedXml(this ReferenceModeValues value)
        {
            switch (value)
            {
                case ReferenceModeValues.R1C1:
                    return XLReferenceStyle.R1C1;

                case ReferenceModeValues.A1:
                    return XLReferenceStyle.A1;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLAlignmentReadingOrderValues ToClosedXml(this uint value)
        {
            switch (value)
            {
                case 0:
                    return XLAlignmentReadingOrderValues.ContextDependent;

                case 1:
                    return XLAlignmentReadingOrderValues.LeftToRight;

                case 2:
                    return XLAlignmentReadingOrderValues.RightToLeft;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLTotalsRowFunction ToClosedXml(this TotalsRowFunctionValues value)
        {
            switch (value)
            {
                case TotalsRowFunctionValues.None:
                    return XLTotalsRowFunction.None;

                case TotalsRowFunctionValues.Sum:
                    return XLTotalsRowFunction.Sum;

                case TotalsRowFunctionValues.Minimum:
                    return XLTotalsRowFunction.Minimum;

                case TotalsRowFunctionValues.Maximum:
                    return XLTotalsRowFunction.Maximum;

                case TotalsRowFunctionValues.Average:
                    return XLTotalsRowFunction.Average;

                case TotalsRowFunctionValues.Count:
                    return XLTotalsRowFunction.Count;

                case TotalsRowFunctionValues.CountNumbers:
                    return XLTotalsRowFunction.CountNumbers;

                case TotalsRowFunctionValues.StandardDeviation:
                    return XLTotalsRowFunction.StandardDeviation;

                case TotalsRowFunctionValues.Variance:
                    return XLTotalsRowFunction.Variance;

                case TotalsRowFunctionValues.Custom:
                    return XLTotalsRowFunction.Custom;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLAllowedValues ToClosedXml(this DataValidationValues value)
        {
            switch (value)
            {
                case DataValidationValues.None:
                    return XLAllowedValues.AnyValue;

                case DataValidationValues.Custom:
                    return XLAllowedValues.Custom;

                case DataValidationValues.Date:
                    return XLAllowedValues.Date;

                case DataValidationValues.Decimal:
                    return XLAllowedValues.Decimal;

                case DataValidationValues.List:
                    return XLAllowedValues.List;

                case DataValidationValues.TextLength:
                    return XLAllowedValues.TextLength;

                case DataValidationValues.Time:
                    return XLAllowedValues.Time;

                case DataValidationValues.Whole:
                    return XLAllowedValues.WholeNumber;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLErrorStyle ToClosedXml(this DataValidationErrorStyleValues value)
        {
            switch (value)
            {
                case DataValidationErrorStyleValues.Information:
                    return XLErrorStyle.Information;

                case DataValidationErrorStyleValues.Warning:
                    return XLErrorStyle.Warning;

                case DataValidationErrorStyleValues.Stop:
                    return XLErrorStyle.Stop;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLOperator ToClosedXml(this DataValidationOperatorValues value)
        {
            switch (value)
            {
                case DataValidationOperatorValues.Between:
                    return XLOperator.Between;

                case DataValidationOperatorValues.GreaterThanOrEqual:
                    return XLOperator.EqualOrGreaterThan;

                case DataValidationOperatorValues.LessThanOrEqual:
                    return XLOperator.EqualOrLessThan;

                case DataValidationOperatorValues.Equal:
                    return XLOperator.EqualTo;

                case DataValidationOperatorValues.GreaterThan:
                    return XLOperator.GreaterThan;

                case DataValidationOperatorValues.LessThan:
                    return XLOperator.LessThan;

                case DataValidationOperatorValues.NotBetween:
                    return XLOperator.NotBetween;

                case DataValidationOperatorValues.NotEqual:
                    return XLOperator.NotEqualTo;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLWorksheetVisibility ToClosedXml(this SheetStateValues value)
        {
            switch (value)
            {
                case SheetStateValues.Visible:
                    return XLWorksheetVisibility.Visible;

                case SheetStateValues.Hidden:
                    return XLWorksheetVisibility.Hidden;

                case SheetStateValues.VeryHidden:
                    return XLWorksheetVisibility.VeryHidden;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPhoneticAlignment ToClosedXml(this PhoneticAlignmentValues value)
        {
            switch (value)
            {
                case PhoneticAlignmentValues.Center:
                    return XLPhoneticAlignment.Center;

                case PhoneticAlignmentValues.Distributed:
                    return XLPhoneticAlignment.Distributed;

                case PhoneticAlignmentValues.Left:
                    return XLPhoneticAlignment.Left;

                case PhoneticAlignmentValues.NoControl:
                    return XLPhoneticAlignment.NoControl;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPhoneticType ToClosedXml(this PhoneticValues value)
        {
            switch (value)
            {
                case PhoneticValues.FullWidthKatakana: return XLPhoneticType.FullWidthKatakana;
                case PhoneticValues.HalfWidthKatakana:
                    return XLPhoneticType.HalfWidthKatakana;

                case PhoneticValues.Hiragana:
                    return XLPhoneticType.Hiragana;

                case PhoneticValues.NoConversion:
                    return XLPhoneticType.NoConversion;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPivotSummary ToClosedXml(this DataConsolidateFunctionValues value)
        {
            switch (value)
            {
                case DataConsolidateFunctionValues.Sum: return XLPivotSummary.Sum;
                case DataConsolidateFunctionValues.Count: return XLPivotSummary.Count;
                case DataConsolidateFunctionValues.Average: return XLPivotSummary.Average;
                case DataConsolidateFunctionValues.Minimum: return XLPivotSummary.Minimum;
                case DataConsolidateFunctionValues.Maximum: return XLPivotSummary.Maximum;
                case DataConsolidateFunctionValues.Product: return XLPivotSummary.Product;
                case DataConsolidateFunctionValues.CountNumbers: return XLPivotSummary.CountNumbers;
                case DataConsolidateFunctionValues.StandardDeviation: return XLPivotSummary.StandardDeviation;
                case DataConsolidateFunctionValues.StandardDeviationP: return XLPivotSummary.PopulationStandardDeviation;
                case DataConsolidateFunctionValues.Variance: return XLPivotSummary.Variance;
                case DataConsolidateFunctionValues.VarianceP: return XLPivotSummary.PopulationVariance;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPivotCalculation ToClosedXml(this ShowDataAsValues value)
        {
            switch (value)
            {
                case ShowDataAsValues.Normal: return XLPivotCalculation.Normal;
                case ShowDataAsValues.Difference: return XLPivotCalculation.DifferenceFrom;
                case ShowDataAsValues.Percent: return XLPivotCalculation.PercentageOf;
                case ShowDataAsValues.PercentageDifference: return XLPivotCalculation.PercentageDifferenceFrom;
                case ShowDataAsValues.RunTotal: return XLPivotCalculation.RunningTotal;
                case ShowDataAsValues.PercentOfRaw: return XLPivotCalculation.PercentageOfRow; // There's a typo in the OpenXML SDK =)
                case ShowDataAsValues.PercentOfColumn: return XLPivotCalculation.PercentageOfColumn;
                case ShowDataAsValues.PercentOfTotal: return XLPivotCalculation.PercentageOfTotal;
                case ShowDataAsValues.Index: return XLPivotCalculation.Index;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLFilterOperator ToClosedXml(this FilterOperatorValues value)
        {
            switch (value)
            {
                case FilterOperatorValues.Equal: return XLFilterOperator.Equal;
                case FilterOperatorValues.NotEqual: return XLFilterOperator.NotEqual;
                case FilterOperatorValues.GreaterThan: return XLFilterOperator.GreaterThan;
                case FilterOperatorValues.LessThan: return XLFilterOperator.LessThan;
                case FilterOperatorValues.GreaterThanOrEqual: return XLFilterOperator.EqualOrGreaterThan;
                case FilterOperatorValues.LessThanOrEqual: return XLFilterOperator.EqualOrLessThan;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLFilterDynamicType ToClosedXml(this DynamicFilterValues value)
        {
            switch (value)
            {
                case DynamicFilterValues.AboveAverage: return XLFilterDynamicType.AboveAverage;
                case DynamicFilterValues.BelowAverage: return XLFilterDynamicType.BelowAverage;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLDateTimeGrouping ToClosedXml(this DateTimeGroupingValues value)
        {
            switch (value)
            {
                case DateTimeGroupingValues.Year: return XLDateTimeGrouping.Year;
                case DateTimeGroupingValues.Month: return XLDateTimeGrouping.Month;
                case DateTimeGroupingValues.Day: return XLDateTimeGrouping.Day;
                case DateTimeGroupingValues.Hour: return XLDateTimeGrouping.Hour;
                case DateTimeGroupingValues.Minute: return XLDateTimeGrouping.Minute;
                case DateTimeGroupingValues.Second: return XLDateTimeGrouping.Second;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLSheetViewOptions ToClosedXml(this SheetViewValues value)
        {
            switch (value)
            {
                case SheetViewValues.Normal: return XLSheetViewOptions.Normal;
                case SheetViewValues.PageBreakPreview: return XLSheetViewOptions.PageBreakPreview;
                case SheetViewValues.PageLayout: return XLSheetViewOptions.PageLayout;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLLineStyle ToClosedXml(this Vml.StrokeLineStyleValues value)
        {
            switch (value)
            {
                case Vml.StrokeLineStyleValues.Single: return XLLineStyle.Single;
                case Vml.StrokeLineStyleValues.ThickBetweenThin: return XLLineStyle.ThickBetweenThin;
                case Vml.StrokeLineStyleValues.ThickThin: return XLLineStyle.ThickThin;
                case Vml.StrokeLineStyleValues.ThinThick: return XLLineStyle.ThinThick;
                case Vml.StrokeLineStyleValues.ThinThin: return XLLineStyle.ThinThin;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLConditionalFormatType ToClosedXml(this ConditionalFormatValues value)
        {
            switch (value)
            {
                case ConditionalFormatValues.Expression: return XLConditionalFormatType.Expression;
                case ConditionalFormatValues.CellIs: return XLConditionalFormatType.CellIs;
                case ConditionalFormatValues.ColorScale: return XLConditionalFormatType.ColorScale;
                case ConditionalFormatValues.DataBar: return XLConditionalFormatType.DataBar;
                case ConditionalFormatValues.IconSet: return XLConditionalFormatType.IconSet;
                case ConditionalFormatValues.Top10: return XLConditionalFormatType.Top10;
                case ConditionalFormatValues.UniqueValues: return XLConditionalFormatType.IsUnique;
                case ConditionalFormatValues.DuplicateValues: return XLConditionalFormatType.IsDuplicate;
                case ConditionalFormatValues.ContainsText: return XLConditionalFormatType.ContainsText;
                case ConditionalFormatValues.NotContainsText: return XLConditionalFormatType.NotContainsText;
                case ConditionalFormatValues.BeginsWith: return XLConditionalFormatType.StartsWith;
                case ConditionalFormatValues.EndsWith: return XLConditionalFormatType.EndsWith;
                case ConditionalFormatValues.ContainsBlanks: return XLConditionalFormatType.IsBlank;
                case ConditionalFormatValues.NotContainsBlanks: return XLConditionalFormatType.NotBlank;
                case ConditionalFormatValues.ContainsErrors: return XLConditionalFormatType.IsError;
                case ConditionalFormatValues.NotContainsErrors: return XLConditionalFormatType.NotError;
                case ConditionalFormatValues.TimePeriod: return XLConditionalFormatType.TimePeriod;
                case ConditionalFormatValues.AboveAverage: return XLConditionalFormatType.AboveAverage;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLCFContentType ToClosedXml(this ConditionalFormatValueObjectValues value)
        {
            switch (value)
            {
                case ConditionalFormatValueObjectValues.Number: return XLCFContentType.Number;
                case ConditionalFormatValueObjectValues.Percent: return XLCFContentType.Percent;
                case ConditionalFormatValueObjectValues.Max: return XLCFContentType.Maximum;
                case ConditionalFormatValueObjectValues.Min: return XLCFContentType.Minimum;
                case ConditionalFormatValueObjectValues.Formula: return XLCFContentType.Formula;
                case ConditionalFormatValueObjectValues.Percentile: return XLCFContentType.Percentile;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLCFOperator ToClosedXml(this ConditionalFormattingOperatorValues value)
        {
            switch (value)
            {
                case ConditionalFormattingOperatorValues.LessThan: return XLCFOperator.LessThan;
                case ConditionalFormattingOperatorValues.LessThanOrEqual: return XLCFOperator.EqualOrLessThan;
                case ConditionalFormattingOperatorValues.Equal: return XLCFOperator.Equal;
                case ConditionalFormattingOperatorValues.NotEqual: return XLCFOperator.NotEqual;
                case ConditionalFormattingOperatorValues.GreaterThanOrEqual: return XLCFOperator.EqualOrGreaterThan;
                case ConditionalFormattingOperatorValues.GreaterThan: return XLCFOperator.GreaterThan;
                case ConditionalFormattingOperatorValues.Between: return XLCFOperator.Between;
                case ConditionalFormattingOperatorValues.NotBetween: return XLCFOperator.NotBetween;
                case ConditionalFormattingOperatorValues.ContainsText: return XLCFOperator.Contains;
                case ConditionalFormattingOperatorValues.NotContains: return XLCFOperator.NotContains;
                case ConditionalFormattingOperatorValues.BeginsWith: return XLCFOperator.StartsWith;
                case ConditionalFormattingOperatorValues.EndsWith: return XLCFOperator.EndsWith;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLIconSetStyle ToClosedXml(this IconSetValues value)
        {
            switch (value)
            {
                case IconSetValues.ThreeArrows: return XLIconSetStyle.ThreeArrows;
                case IconSetValues.ThreeArrowsGray: return XLIconSetStyle.ThreeArrowsGray;
                case IconSetValues.ThreeFlags: return XLIconSetStyle.ThreeFlags;
                case IconSetValues.ThreeTrafficLights1: return XLIconSetStyle.ThreeTrafficLights1;
                case IconSetValues.ThreeTrafficLights2: return XLIconSetStyle.ThreeTrafficLights2;
                case IconSetValues.ThreeSigns: return XLIconSetStyle.ThreeSigns;
                case IconSetValues.ThreeSymbols: return XLIconSetStyle.ThreeSymbols;
                case IconSetValues.ThreeSymbols2: return XLIconSetStyle.ThreeSymbols2;
                case IconSetValues.FourArrows: return XLIconSetStyle.FourArrows;
                case IconSetValues.FourArrowsGray: return XLIconSetStyle.FourArrowsGray;
                case IconSetValues.FourRedToBlack: return XLIconSetStyle.FourRedToBlack;
                case IconSetValues.FourRating: return XLIconSetStyle.FourRating;
                case IconSetValues.FourTrafficLights: return XLIconSetStyle.FourTrafficLights;
                case IconSetValues.FiveArrows: return XLIconSetStyle.FiveArrows;
                case IconSetValues.FiveArrowsGray: return XLIconSetStyle.FiveArrowsGray;
                case IconSetValues.FiveRating: return XLIconSetStyle.FiveRating;
                case IconSetValues.FiveQuarters: return XLIconSetStyle.FiveQuarters;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLTimePeriod ToClosedXml(this TimePeriodValues value)
        {
            switch (value)
            {
                case TimePeriodValues.Yesterday: return XLTimePeriod.Yesterday;
                case TimePeriodValues.Today: return XLTimePeriod.Today;
                case TimePeriodValues.Tomorrow: return XLTimePeriod.Tomorrow;
                case TimePeriodValues.Last7Days: return XLTimePeriod.InTheLast7Days;
                case TimePeriodValues.LastWeek: return XLTimePeriod.LastWeek;
                case TimePeriodValues.ThisWeek: return XLTimePeriod.ThisWeek;
                case TimePeriodValues.NextWeek: return XLTimePeriod.NextWeek;
                case TimePeriodValues.LastMonth: return XLTimePeriod.LastMonth;
                case TimePeriodValues.ThisMonth: return XLTimePeriod.ThisMonth;
                case TimePeriodValues.NextMonth: return XLTimePeriod.NextMonth;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPictureFormat ToClosedXml(this ImagePartType value)
        {
            return Enum.Parse(typeof(XLPictureFormat), value.ToString()).CastTo<XLPictureFormat>();
        }

        public static XLPicturePlacement ToClosedXml(this Xdr.EditAsValues value)
        {
            switch (value)
            {
                case Xdr.EditAsValues.Absolute:
                    return XLPicturePlacement.FreeFloating;

                case Xdr.EditAsValues.OneCell:
                    return XLPicturePlacement.Move;

                case Xdr.EditAsValues.TwoCell:
                    return XLPicturePlacement.MoveAndSize;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLPivotAreaValues ToClosedXml(this PivotAreaValues value)
        {
            switch (value)
            {
                case PivotAreaValues.None:
                    return XLPivotAreaValues.None;

                case PivotAreaValues.Normal:
                    return XLPivotAreaValues.Normal;

                case PivotAreaValues.Data:
                    return XLPivotAreaValues.Data;

                case PivotAreaValues.All:
                    return XLPivotAreaValues.All;

                case PivotAreaValues.Origin:
                    return XLPivotAreaValues.Origin;

                case PivotAreaValues.Button:
                    return XLPivotAreaValues.Button;

                case PivotAreaValues.TopRight:
                    return XLPivotAreaValues.TopRight;

                case PivotAreaValues.TopEnd:
                    return XLPivotAreaValues.TopEnd;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "PivotAreaValues value not implemented");
            }
        }

        public static XLSparklineType ToClosedXml(this X14.SparklineTypeValues value)
        {
            switch (value)
            {
                case X14.SparklineTypeValues.Line: return XLSparklineType.Line;
                case X14.SparklineTypeValues.Column: return XLSparklineType.Column;
                case X14.SparklineTypeValues.Stacked: return XLSparklineType.Stacked;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLSparklineAxisMinMax ToClosedXml(this X14.SparklineAxisMinMaxValues value)
        {
            switch (value)
            {
                case X14.SparklineAxisMinMaxValues.Individual: return XLSparklineAxisMinMax.Automatic;
                case X14.SparklineAxisMinMaxValues.Group: return XLSparklineAxisMinMax.SameForAll;
                case X14.SparklineAxisMinMaxValues.Custom: return XLSparklineAxisMinMax.Custom;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static XLDisplayBlanksAsValues ToClosedXml(this X14.DisplayBlanksAsValues value)
        {
            switch (value)
            {
                case X14.DisplayBlanksAsValues.Span: return XLDisplayBlanksAsValues.Interpolate;
                case X14.DisplayBlanksAsValues.Gap: return XLDisplayBlanksAsValues.NotPlotted;
                case X14.DisplayBlanksAsValues.Zero: return XLDisplayBlanksAsValues.Zero;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        #endregion To ClosedXml
    }
}
