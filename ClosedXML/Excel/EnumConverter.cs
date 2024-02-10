#nullable disable

using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
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

        private static readonly String[] XLFontUnderlineValuesStrings =
        {
            "double",
            "doubleAccounting",
            "none",
            "single",
            "singleAccounting"
        };

        public static string ToOpenXmlString(this XLFontUnderlineValues value)
            => XLFontUnderlineValuesStrings[(int)value];

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

        private static readonly String[] XLFontVerticalTextAlignmentValuesStrings =
        {
            "baseline",
            "subscript",
            "superscript"
        };

        public static String ToOpenXmlString(this XLFontVerticalTextAlignmentValues value)
            => XLFontVerticalTextAlignmentValuesStrings[(int)value];

        private static readonly String[] XLFontSchemeStrings =
        {
            "none",
            "major",
            "minor"
        };

        public static String ToOpenXml(this XLFontScheme value)
            => XLFontSchemeStrings[(int)value];

        public static FontSchemeValues ToOpenXmlEnum(this XLFontScheme value)
        {
            return value switch
            {
                XLFontScheme.None => FontSchemeValues.None,
                XLFontScheme.Major => FontSchemeValues.Major,
                XLFontScheme.Minor => FontSchemeValues.Minor,
                _ => throw new ArgumentOutOfRangeException()
            };
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

        private static readonly String[] XLPhoneticAlignmentStrings =
        {
            "center",
            "distributed",
            "left",
            "noControl"
        };

        public static String ToOpenXmlString(this XLPhoneticAlignment value)
            => XLPhoneticAlignmentStrings[(int)value];

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

        private static readonly String[] XLPhoneticTypeStrings =
        {
            "fullwidthKatakana",
            "halfwidthKatakana",
            "Hiragana",
            "noConversion"
        };

        public static String ToOpenXmlString(this XLPhoneticType value)
            => XLPhoneticTypeStrings[(int)value];

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
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
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

        private static readonly IReadOnlyDictionary<XLPictureFormat, PartTypeInfo> PictureFormatMap =
            new Dictionary<XLPictureFormat, PartTypeInfo>
            {
                { XLPictureFormat.Unknown, new PartTypeInfo("image/unknown", ".bin") },
                { XLPictureFormat.Bmp, ImagePartType.Bmp },
                { XLPictureFormat.Gif, ImagePartType.Gif },
                { XLPictureFormat.Png, ImagePartType.Png },
                { XLPictureFormat.Tiff, ImagePartType.Tiff },
                { XLPictureFormat.Icon, ImagePartType.Icon },
                { XLPictureFormat.Pcx, ImagePartType.Pcx },
                { XLPictureFormat.Jpeg, ImagePartType.Jpeg },
                { XLPictureFormat.Emf, ImagePartType.Emf },
                { XLPictureFormat.Wmf, ImagePartType.Wmf },
                { XLPictureFormat.Webp, new PartTypeInfo("image/webp", ".webp") }
            };

        public static PartTypeInfo ToOpenXml(this XLPictureFormat value)
        {
            return PictureFormatMap[value];
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
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PivotAreaValues ToOpenXml(this XLPivotAreaType value)
        {
            switch (value)
            {
                case XLPivotAreaType.None:
                    return PivotAreaValues.None;

                case XLPivotAreaType.Normal:
                    return PivotAreaValues.Normal;

                case XLPivotAreaType.Data:
                    return PivotAreaValues.Data;

                case XLPivotAreaType.All:
                    return PivotAreaValues.All;

                case XLPivotAreaType.Origin:
                    return PivotAreaValues.Origin;

                case XLPivotAreaType.Button:
                    return PivotAreaValues.Button;

                case XLPivotAreaType.TopRight:
                    return PivotAreaValues.TopRight;

                case XLPivotAreaType.TopEnd:
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

        public static FieldSortValues ToOpenXml(this XLPivotSortType value)
        {
            switch (value)
            {
                case XLPivotSortType.Default: return FieldSortValues.Manual;
                case XLPivotSortType.Ascending: return FieldSortValues.Ascending;
                case XLPivotSortType.Descending: return FieldSortValues.Descending;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        #endregion To OpenXml

        #region To ClosedXml

        private static readonly IReadOnlyDictionary<UnderlineValues, XLFontUnderlineValues> UnderlineValuesMap =
            new Dictionary<UnderlineValues, XLFontUnderlineValues>
            {
                { UnderlineValues.Double, XLFontUnderlineValues.Double },
                { UnderlineValues.DoubleAccounting, XLFontUnderlineValues.DoubleAccounting },
                { UnderlineValues.None, XLFontUnderlineValues.None },
                { UnderlineValues.Single, XLFontUnderlineValues.Single },
                { UnderlineValues.SingleAccounting, XLFontUnderlineValues.SingleAccounting },
            };

        public static XLFontUnderlineValues ToClosedXml(this UnderlineValues value)
        {
            return UnderlineValuesMap[value];
        }

        private static readonly IReadOnlyDictionary<FontSchemeValues, XLFontScheme> FontSchemeMap =
            new Dictionary<FontSchemeValues, XLFontScheme>
            {
                { FontSchemeValues.None, XLFontScheme.None },
                { FontSchemeValues.Major, XLFontScheme.Major },
                { FontSchemeValues.Minor, XLFontScheme.Minor },
            };

        public static XLFontScheme ToClosedXml(this FontSchemeValues value)
        {
            return FontSchemeMap[value];
        }

        private static readonly IReadOnlyDictionary<OrientationValues, XLPageOrientation> OrientationMap =
            new Dictionary<OrientationValues, XLPageOrientation>
            {
                { OrientationValues.Default, XLPageOrientation.Default },
                { OrientationValues.Landscape, XLPageOrientation.Landscape },
                { OrientationValues.Portrait, XLPageOrientation.Portrait },
            };

        public static XLPageOrientation ToClosedXml(this OrientationValues value)
        {
            return OrientationMap[value];
        }

        private static readonly IReadOnlyDictionary<VerticalAlignmentRunValues, XLFontVerticalTextAlignmentValues> VerticalAlignmentRunMap =
            new Dictionary<VerticalAlignmentRunValues, XLFontVerticalTextAlignmentValues>
            {
                { VerticalAlignmentRunValues.Baseline, XLFontVerticalTextAlignmentValues.Baseline },
                { VerticalAlignmentRunValues.Subscript, XLFontVerticalTextAlignmentValues.Subscript },
                { VerticalAlignmentRunValues.Superscript, XLFontVerticalTextAlignmentValues.Superscript },
            };


        public static XLFontVerticalTextAlignmentValues ToClosedXml(this VerticalAlignmentRunValues value)
        {
            return VerticalAlignmentRunMap[value];
        }

        private static readonly IReadOnlyDictionary<PatternValues, XLFillPatternValues> PatternMap =
            new Dictionary<PatternValues, XLFillPatternValues>
            {
                { PatternValues.DarkDown, XLFillPatternValues.DarkDown },
                { PatternValues.DarkGray, XLFillPatternValues.DarkGray },
                { PatternValues.DarkGrid, XLFillPatternValues.DarkGrid },
                { PatternValues.DarkHorizontal, XLFillPatternValues.DarkHorizontal },
                { PatternValues.DarkTrellis, XLFillPatternValues.DarkTrellis },
                { PatternValues.DarkUp, XLFillPatternValues.DarkUp },
                { PatternValues.DarkVertical, XLFillPatternValues.DarkVertical },
                { PatternValues.Gray0625, XLFillPatternValues.Gray0625 },
                { PatternValues.Gray125, XLFillPatternValues.Gray125 },
                { PatternValues.LightDown, XLFillPatternValues.LightDown },
                { PatternValues.LightGray, XLFillPatternValues.LightGray },
                { PatternValues.LightGrid, XLFillPatternValues.LightGrid },
                { PatternValues.LightHorizontal, XLFillPatternValues.LightHorizontal },
                { PatternValues.LightTrellis, XLFillPatternValues.LightTrellis },
                { PatternValues.LightUp, XLFillPatternValues.LightUp },
                { PatternValues.LightVertical, XLFillPatternValues.LightVertical },
                { PatternValues.MediumGray, XLFillPatternValues.MediumGray },
                { PatternValues.None, XLFillPatternValues.None },
                { PatternValues.Solid, XLFillPatternValues.Solid },
            };

        public static XLFillPatternValues ToClosedXml(this PatternValues value)
        {
            return PatternMap[value];
        }

        private static readonly IReadOnlyDictionary<BorderStyleValues, XLBorderStyleValues> BorderStyleMap =
            new Dictionary<BorderStyleValues, XLBorderStyleValues>
            {
                { BorderStyleValues.DashDot, XLBorderStyleValues.DashDot },
                { BorderStyleValues.DashDotDot, XLBorderStyleValues.DashDotDot },
                { BorderStyleValues.Dashed, XLBorderStyleValues.Dashed },
                { BorderStyleValues.Dotted, XLBorderStyleValues.Dotted },
                { BorderStyleValues.Double, XLBorderStyleValues.Double },
                { BorderStyleValues.Hair, XLBorderStyleValues.Hair },
                { BorderStyleValues.Medium, XLBorderStyleValues.Medium },
                { BorderStyleValues.MediumDashDot, XLBorderStyleValues.MediumDashDot },
                { BorderStyleValues.MediumDashDotDot, XLBorderStyleValues.MediumDashDotDot },
                { BorderStyleValues.MediumDashed, XLBorderStyleValues.MediumDashed },
                { BorderStyleValues.None, XLBorderStyleValues.None },
                { BorderStyleValues.SlantDashDot, XLBorderStyleValues.SlantDashDot },
                { BorderStyleValues.Thick, XLBorderStyleValues.Thick },
                { BorderStyleValues.Thin, XLBorderStyleValues.Thin },
            };

        public static XLBorderStyleValues ToClosedXml(this BorderStyleValues value)
        {
            return BorderStyleMap[value];
        }

        private static readonly IReadOnlyDictionary<HorizontalAlignmentValues, XLAlignmentHorizontalValues> HorizontalAlignmentMap =
            new Dictionary<HorizontalAlignmentValues, XLAlignmentHorizontalValues>
            {
                { HorizontalAlignmentValues.Center, XLAlignmentHorizontalValues.Center },
                { HorizontalAlignmentValues.CenterContinuous, XLAlignmentHorizontalValues.CenterContinuous },
                { HorizontalAlignmentValues.Distributed, XLAlignmentHorizontalValues.Distributed },
                { HorizontalAlignmentValues.Fill, XLAlignmentHorizontalValues.Fill },
                { HorizontalAlignmentValues.General, XLAlignmentHorizontalValues.General },
                { HorizontalAlignmentValues.Justify, XLAlignmentHorizontalValues.Justify },
                { HorizontalAlignmentValues.Left, XLAlignmentHorizontalValues.Left },
                { HorizontalAlignmentValues.Right, XLAlignmentHorizontalValues.Right },
            };

        public static XLAlignmentHorizontalValues ToClosedXml(this HorizontalAlignmentValues value)
        {
            return HorizontalAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<VerticalAlignmentValues, XLAlignmentVerticalValues> VerticalAlignmentMap =
            new Dictionary<VerticalAlignmentValues, XLAlignmentVerticalValues>
            {
                { VerticalAlignmentValues.Bottom, XLAlignmentVerticalValues.Bottom },
                { VerticalAlignmentValues.Center, XLAlignmentVerticalValues.Center },
                { VerticalAlignmentValues.Distributed, XLAlignmentVerticalValues.Distributed },
                { VerticalAlignmentValues.Justify, XLAlignmentVerticalValues.Justify },
                { VerticalAlignmentValues.Top, XLAlignmentVerticalValues.Top },
            };

        public static XLAlignmentVerticalValues ToClosedXml(this VerticalAlignmentValues value)
        {
            return VerticalAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<PageOrderValues, XLPageOrderValues> PageOrdersMap =
            new Dictionary<PageOrderValues, XLPageOrderValues>
            {
                { PageOrderValues.DownThenOver, XLPageOrderValues.DownThenOver },
                { PageOrderValues.OverThenDown, XLPageOrderValues.OverThenDown },
            };

        public static XLPageOrderValues ToClosedXml(this PageOrderValues value)
        {
            return PageOrdersMap[value];
        }

        private static readonly IReadOnlyDictionary<CellCommentsValues, XLShowCommentsValues> CellCommentsMap =
            new Dictionary<CellCommentsValues, XLShowCommentsValues>
            {
                { CellCommentsValues.AsDisplayed, XLShowCommentsValues.AsDisplayed },
                { CellCommentsValues.AtEnd, XLShowCommentsValues.AtEnd },
                { CellCommentsValues.None, XLShowCommentsValues.None },
            };

        public static XLShowCommentsValues ToClosedXml(this CellCommentsValues value)
        {
            return CellCommentsMap[value];
        }

        private static readonly IReadOnlyDictionary<PrintErrorValues, XLPrintErrorValues> PrintErrorMap =
            new Dictionary<PrintErrorValues, XLPrintErrorValues>
            {
                { PrintErrorValues.Blank, XLPrintErrorValues.Blank },
                { PrintErrorValues.Dash, XLPrintErrorValues.Dash },
                { PrintErrorValues.Displayed, XLPrintErrorValues.Displayed },
                { PrintErrorValues.NA, XLPrintErrorValues.NA },
            };

        public static XLPrintErrorValues ToClosedXml(this PrintErrorValues value)
        {
            return PrintErrorMap[value];
        }

        private static readonly IReadOnlyDictionary<CalculateModeValues, XLCalculateMode> CalculateModeMap =
            new Dictionary<CalculateModeValues, XLCalculateMode>
            {
                { CalculateModeValues.Auto, XLCalculateMode.Auto },
                { CalculateModeValues.AutoNoTable, XLCalculateMode.AutoNoTable },
                { CalculateModeValues.Manual, XLCalculateMode.Manual },
            };

        public static XLCalculateMode ToClosedXml(this CalculateModeValues value)
        {
            return CalculateModeMap[value];
        }

        private static readonly IReadOnlyDictionary<ReferenceModeValues, XLReferenceStyle> ReferenceModeMap =
            new Dictionary<ReferenceModeValues, XLReferenceStyle>
            {
                { ReferenceModeValues.R1C1, XLReferenceStyle.R1C1 },
                { ReferenceModeValues.A1, XLReferenceStyle.A1 },
            };

        public static XLReferenceStyle ToClosedXml(this ReferenceModeValues value)
        {
            return ReferenceModeMap[value];
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

        private static readonly IReadOnlyDictionary<TotalsRowFunctionValues, XLTotalsRowFunction> TotalsRowFunctionMap =
            new Dictionary<TotalsRowFunctionValues, XLTotalsRowFunction>
            {
                { TotalsRowFunctionValues.None, XLTotalsRowFunction.None },
                { TotalsRowFunctionValues.Sum, XLTotalsRowFunction.Sum },
                { TotalsRowFunctionValues.Minimum, XLTotalsRowFunction.Minimum },
                { TotalsRowFunctionValues.Maximum, XLTotalsRowFunction.Maximum },
                { TotalsRowFunctionValues.Average, XLTotalsRowFunction.Average },
                { TotalsRowFunctionValues.Count, XLTotalsRowFunction.Count },
                { TotalsRowFunctionValues.CountNumbers, XLTotalsRowFunction.CountNumbers },
                { TotalsRowFunctionValues.StandardDeviation, XLTotalsRowFunction.StandardDeviation },
                { TotalsRowFunctionValues.Variance, XLTotalsRowFunction.Variance },
                { TotalsRowFunctionValues.Custom, XLTotalsRowFunction.Custom },
            };


        public static XLTotalsRowFunction ToClosedXml(this TotalsRowFunctionValues value)
        {
            return TotalsRowFunctionMap[value];
        }

        private static readonly IReadOnlyDictionary<DataValidationValues, XLAllowedValues> DataValidationMap =
            new Dictionary<DataValidationValues, XLAllowedValues>
            {
                { DataValidationValues.None, XLAllowedValues.AnyValue },
                { DataValidationValues.Custom, XLAllowedValues.Custom },
                { DataValidationValues.Date, XLAllowedValues.Date },
                { DataValidationValues.Decimal, XLAllowedValues.Decimal },
                { DataValidationValues.List, XLAllowedValues.List },
                { DataValidationValues.TextLength, XLAllowedValues.TextLength },
                { DataValidationValues.Time, XLAllowedValues.Time },
                { DataValidationValues.Whole, XLAllowedValues.WholeNumber },
            };

        public static XLAllowedValues ToClosedXml(this DataValidationValues value)
        {
            return DataValidationMap[value];
        }

        private static readonly IReadOnlyDictionary<DataValidationErrorStyleValues, XLErrorStyle> DataValidationErrorStyleMap =
            new Dictionary<DataValidationErrorStyleValues, XLErrorStyle>
            {
                { DataValidationErrorStyleValues.Information, XLErrorStyle.Information },
                { DataValidationErrorStyleValues.Warning, XLErrorStyle.Warning },
                { DataValidationErrorStyleValues.Stop, XLErrorStyle.Stop },
            };

        public static XLErrorStyle ToClosedXml(this DataValidationErrorStyleValues value)
        {
            return DataValidationErrorStyleMap[value];
        }

        private static readonly IReadOnlyDictionary<DataValidationOperatorValues, XLOperator> DataValidationOperatorMap =
            new Dictionary<DataValidationOperatorValues, XLOperator>
            {
                { DataValidationOperatorValues.Between, XLOperator.Between },
                { DataValidationOperatorValues.GreaterThanOrEqual, XLOperator.EqualOrGreaterThan },
                { DataValidationOperatorValues.LessThanOrEqual, XLOperator.EqualOrLessThan },
                { DataValidationOperatorValues.Equal, XLOperator.EqualTo },
                { DataValidationOperatorValues.GreaterThan, XLOperator.GreaterThan },
                { DataValidationOperatorValues.LessThan, XLOperator.LessThan },
                { DataValidationOperatorValues.NotBetween, XLOperator.NotBetween },
                { DataValidationOperatorValues.NotEqual, XLOperator.NotEqualTo },
            };

        public static XLOperator ToClosedXml(this DataValidationOperatorValues value)
        {
            return DataValidationOperatorMap[value];
        }

        private static readonly IReadOnlyDictionary<SheetStateValues, XLWorksheetVisibility> SheetStateMap =
            new Dictionary<SheetStateValues, XLWorksheetVisibility>
            {
                { SheetStateValues.Visible, XLWorksheetVisibility.Visible },
                { SheetStateValues.Hidden, XLWorksheetVisibility.Hidden },
                { SheetStateValues.VeryHidden, XLWorksheetVisibility.VeryHidden },
            };

        public static XLWorksheetVisibility ToClosedXml(this SheetStateValues value)
        {
            return SheetStateMap[value];
        }

        private static readonly IReadOnlyDictionary<PhoneticAlignmentValues, XLPhoneticAlignment> PhoneticAlignmentMap =
            new Dictionary<PhoneticAlignmentValues, XLPhoneticAlignment>
            {
                { PhoneticAlignmentValues.Center, XLPhoneticAlignment.Center },
                { PhoneticAlignmentValues.Distributed, XLPhoneticAlignment.Distributed },
                { PhoneticAlignmentValues.Left, XLPhoneticAlignment.Left },
                { PhoneticAlignmentValues.NoControl, XLPhoneticAlignment.NoControl },
            };


        public static XLPhoneticAlignment ToClosedXml(this PhoneticAlignmentValues value)
        {
            return PhoneticAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<PhoneticValues, XLPhoneticType> PhoneticMap =
            new Dictionary<PhoneticValues, XLPhoneticType>
            {
                { PhoneticValues.FullWidthKatakana, XLPhoneticType.FullWidthKatakana },
                { PhoneticValues.HalfWidthKatakana, XLPhoneticType.HalfWidthKatakana },
                { PhoneticValues.Hiragana, XLPhoneticType.Hiragana },
                { PhoneticValues.NoConversion, XLPhoneticType.NoConversion },
            };

        public static XLPhoneticType ToClosedXml(this PhoneticValues value)
        {
            return PhoneticMap[value];
        }

        private static readonly IReadOnlyDictionary<DataConsolidateFunctionValues, XLPivotSummary> DataConsolidateFunctionMap =
            new Dictionary<DataConsolidateFunctionValues, XLPivotSummary>
            {
                { DataConsolidateFunctionValues.Sum, XLPivotSummary.Sum },
                { DataConsolidateFunctionValues.Count, XLPivotSummary.Count },
                { DataConsolidateFunctionValues.Average, XLPivotSummary.Average },
                { DataConsolidateFunctionValues.Minimum, XLPivotSummary.Minimum },
                { DataConsolidateFunctionValues.Maximum, XLPivotSummary.Maximum },
                { DataConsolidateFunctionValues.Product, XLPivotSummary.Product },
                { DataConsolidateFunctionValues.CountNumbers, XLPivotSummary.CountNumbers },
                { DataConsolidateFunctionValues.StandardDeviation, XLPivotSummary.StandardDeviation },
                { DataConsolidateFunctionValues.StandardDeviationP, XLPivotSummary.PopulationStandardDeviation },
                { DataConsolidateFunctionValues.Variance, XLPivotSummary.Variance },
                { DataConsolidateFunctionValues.VarianceP, XLPivotSummary.PopulationVariance },

            };

        public static XLPivotSummary ToClosedXml(this DataConsolidateFunctionValues value)
        {
            return DataConsolidateFunctionMap[value];
        }

        private static readonly IReadOnlyDictionary<ShowDataAsValues, XLPivotCalculation> ShowDataAsMap =
            new Dictionary<ShowDataAsValues, XLPivotCalculation>
            {
                { ShowDataAsValues.Normal, XLPivotCalculation.Normal },
                { ShowDataAsValues.Difference, XLPivotCalculation.DifferenceFrom },
                { ShowDataAsValues.Percent, XLPivotCalculation.PercentageOf },
                { ShowDataAsValues.PercentageDifference, XLPivotCalculation.PercentageDifferenceFrom },
                { ShowDataAsValues.RunTotal, XLPivotCalculation.RunningTotal },
                { ShowDataAsValues.PercentOfRaw, XLPivotCalculation.PercentageOfRow }, // There's a typo in the OpenXML SDK =)
                { ShowDataAsValues.PercentOfColumn, XLPivotCalculation.PercentageOfColumn },
                { ShowDataAsValues.PercentOfTotal, XLPivotCalculation.PercentageOfTotal },
                { ShowDataAsValues.Index, XLPivotCalculation.Index },
            };

        public static XLPivotCalculation ToClosedXml(this ShowDataAsValues value)
        {
            return ShowDataAsMap[value];
        }

        private static readonly IReadOnlyDictionary<FilterOperatorValues, XLFilterOperator> FilterOperatorMap =
            new Dictionary<FilterOperatorValues, XLFilterOperator>
            {
                { FilterOperatorValues.Equal, XLFilterOperator.Equal },
                { FilterOperatorValues.NotEqual, XLFilterOperator.NotEqual },
                { FilterOperatorValues.GreaterThan, XLFilterOperator.GreaterThan },
                { FilterOperatorValues.LessThan, XLFilterOperator.LessThan },
                { FilterOperatorValues.GreaterThanOrEqual, XLFilterOperator.EqualOrGreaterThan },
                { FilterOperatorValues.LessThanOrEqual, XLFilterOperator.EqualOrLessThan },
            };

        public static XLFilterOperator ToClosedXml(this FilterOperatorValues value)
        {
            return FilterOperatorMap[value];
        }

        private static readonly IReadOnlyDictionary<DynamicFilterValues, XLFilterDynamicType> DynamicFilterMap =
            new Dictionary<DynamicFilterValues, XLFilterDynamicType>
            {
                { DynamicFilterValues.AboveAverage, XLFilterDynamicType.AboveAverage },
                { DynamicFilterValues.BelowAverage, XLFilterDynamicType.BelowAverage },
            };

        public static XLFilterDynamicType ToClosedXml(this DynamicFilterValues value)
        {
            return DynamicFilterMap[value];
        }

        private static readonly IReadOnlyDictionary<DateTimeGroupingValues, XLDateTimeGrouping> DateTimeGroupingMap =
            new Dictionary<DateTimeGroupingValues, XLDateTimeGrouping>
            {
                { DateTimeGroupingValues.Year, XLDateTimeGrouping.Year },
                { DateTimeGroupingValues.Month, XLDateTimeGrouping.Month },
                { DateTimeGroupingValues.Day, XLDateTimeGrouping.Day },
                { DateTimeGroupingValues.Hour, XLDateTimeGrouping.Hour },
                { DateTimeGroupingValues.Minute, XLDateTimeGrouping.Minute },
                { DateTimeGroupingValues.Second, XLDateTimeGrouping.Second },
            };

        public static XLDateTimeGrouping ToClosedXml(this DateTimeGroupingValues value)
        {
            return DateTimeGroupingMap[value];
        }

        private static readonly IReadOnlyDictionary<SheetViewValues, XLSheetViewOptions> SheetViewMap =
            new Dictionary<SheetViewValues, XLSheetViewOptions>
            {
                { SheetViewValues.Normal, XLSheetViewOptions.Normal },
                { SheetViewValues.PageBreakPreview, XLSheetViewOptions.PageBreakPreview },
                { SheetViewValues.PageLayout, XLSheetViewOptions.PageLayout },
            };

        public static XLSheetViewOptions ToClosedXml(this SheetViewValues value)
        {
            return SheetViewMap[value];
        }

        private static readonly IReadOnlyDictionary<Vml.StrokeLineStyleValues, XLLineStyle> StrokeLineStyleMap =
            new Dictionary<Vml.StrokeLineStyleValues, XLLineStyle>
            {
                { Vml.StrokeLineStyleValues.Single, XLLineStyle.Single },
                { Vml.StrokeLineStyleValues.ThickBetweenThin, XLLineStyle.ThickBetweenThin },
                { Vml.StrokeLineStyleValues.ThickThin, XLLineStyle.ThickThin },
                { Vml.StrokeLineStyleValues.ThinThick, XLLineStyle.ThinThick },
                { Vml.StrokeLineStyleValues.ThinThin, XLLineStyle.ThinThin },
            };

        public static XLLineStyle ToClosedXml(this Vml.StrokeLineStyleValues value)
        {
            return StrokeLineStyleMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormatValues, XLConditionalFormatType> ConditionalFormatMap =
            new Dictionary<ConditionalFormatValues, XLConditionalFormatType>
            {
                { ConditionalFormatValues.Expression, XLConditionalFormatType.Expression },
                { ConditionalFormatValues.CellIs, XLConditionalFormatType.CellIs },
                { ConditionalFormatValues.ColorScale, XLConditionalFormatType.ColorScale },
                { ConditionalFormatValues.DataBar, XLConditionalFormatType.DataBar },
                { ConditionalFormatValues.IconSet, XLConditionalFormatType.IconSet },
                { ConditionalFormatValues.Top10, XLConditionalFormatType.Top10 },
                { ConditionalFormatValues.UniqueValues, XLConditionalFormatType.IsUnique },
                { ConditionalFormatValues.DuplicateValues, XLConditionalFormatType.IsDuplicate },
                { ConditionalFormatValues.ContainsText, XLConditionalFormatType.ContainsText },
                { ConditionalFormatValues.NotContainsText, XLConditionalFormatType.NotContainsText },
                { ConditionalFormatValues.BeginsWith, XLConditionalFormatType.StartsWith },
                { ConditionalFormatValues.EndsWith, XLConditionalFormatType.EndsWith },
                { ConditionalFormatValues.ContainsBlanks, XLConditionalFormatType.IsBlank },
                { ConditionalFormatValues.NotContainsBlanks, XLConditionalFormatType.NotBlank },
                { ConditionalFormatValues.ContainsErrors, XLConditionalFormatType.IsError },
                { ConditionalFormatValues.NotContainsErrors, XLConditionalFormatType.NotError },
                { ConditionalFormatValues.TimePeriod, XLConditionalFormatType.TimePeriod },
                { ConditionalFormatValues.AboveAverage, XLConditionalFormatType.AboveAverage },
            };

        public static XLConditionalFormatType ToClosedXml(this ConditionalFormatValues value)
        {
            return ConditionalFormatMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormatValueObjectValues, XLCFContentType> ConditionalFormatValueObjectMap =
            new Dictionary<ConditionalFormatValueObjectValues, XLCFContentType>
            {
                { ConditionalFormatValueObjectValues.Number, XLCFContentType.Number },
                { ConditionalFormatValueObjectValues.Percent, XLCFContentType.Percent },
                { ConditionalFormatValueObjectValues.Max, XLCFContentType.Maximum },
                { ConditionalFormatValueObjectValues.Min, XLCFContentType.Minimum },
                { ConditionalFormatValueObjectValues.Formula, XLCFContentType.Formula },
                { ConditionalFormatValueObjectValues.Percentile, XLCFContentType.Percentile },
            };

        public static XLCFContentType ToClosedXml(this ConditionalFormatValueObjectValues value)
        {
            return ConditionalFormatValueObjectMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormattingOperatorValues, XLCFOperator> ConditionalFormattingOperatorMap =
            new Dictionary<ConditionalFormattingOperatorValues, XLCFOperator>
            {
                { ConditionalFormattingOperatorValues.LessThan, XLCFOperator.LessThan },
                { ConditionalFormattingOperatorValues.LessThanOrEqual, XLCFOperator.EqualOrLessThan },
                { ConditionalFormattingOperatorValues.Equal, XLCFOperator.Equal },
                { ConditionalFormattingOperatorValues.NotEqual, XLCFOperator.NotEqual },
                { ConditionalFormattingOperatorValues.GreaterThanOrEqual, XLCFOperator.EqualOrGreaterThan },
                { ConditionalFormattingOperatorValues.GreaterThan, XLCFOperator.GreaterThan },
                { ConditionalFormattingOperatorValues.Between, XLCFOperator.Between },
                { ConditionalFormattingOperatorValues.NotBetween, XLCFOperator.NotBetween },
                { ConditionalFormattingOperatorValues.ContainsText, XLCFOperator.Contains },
                { ConditionalFormattingOperatorValues.NotContains, XLCFOperator.NotContains },
                { ConditionalFormattingOperatorValues.BeginsWith, XLCFOperator.StartsWith },
                { ConditionalFormattingOperatorValues.EndsWith, XLCFOperator.EndsWith },
            };

        public static XLCFOperator ToClosedXml(this ConditionalFormattingOperatorValues value)
        {
            return ConditionalFormattingOperatorMap[value];
        }

        private static readonly IReadOnlyDictionary<IconSetValues, XLIconSetStyle> IconSetMap =
            new Dictionary<IconSetValues, XLIconSetStyle>
            {
                { IconSetValues.ThreeArrows, XLIconSetStyle.ThreeArrows },
                { IconSetValues.ThreeArrowsGray, XLIconSetStyle.ThreeArrowsGray },
                { IconSetValues.ThreeFlags, XLIconSetStyle.ThreeFlags },
                { IconSetValues.ThreeTrafficLights1, XLIconSetStyle.ThreeTrafficLights1 },
                { IconSetValues.ThreeTrafficLights2, XLIconSetStyle.ThreeTrafficLights2 },
                { IconSetValues.ThreeSigns, XLIconSetStyle.ThreeSigns },
                { IconSetValues.ThreeSymbols, XLIconSetStyle.ThreeSymbols },
                { IconSetValues.ThreeSymbols2, XLIconSetStyle.ThreeSymbols2 },
                { IconSetValues.FourArrows, XLIconSetStyle.FourArrows },
                { IconSetValues.FourArrowsGray, XLIconSetStyle.FourArrowsGray },
                { IconSetValues.FourRedToBlack, XLIconSetStyle.FourRedToBlack },
                { IconSetValues.FourRating, XLIconSetStyle.FourRating },
                { IconSetValues.FourTrafficLights, XLIconSetStyle.FourTrafficLights },
                { IconSetValues.FiveArrows, XLIconSetStyle.FiveArrows },
                { IconSetValues.FiveArrowsGray, XLIconSetStyle.FiveArrowsGray },
                { IconSetValues.FiveRating, XLIconSetStyle.FiveRating },
                { IconSetValues.FiveQuarters, XLIconSetStyle.FiveQuarters },
            };

        public static XLIconSetStyle ToClosedXml(this IconSetValues value)
        {
            return IconSetMap[value];
        }

        private static readonly IReadOnlyDictionary<TimePeriodValues, XLTimePeriod> TimePeriodMap =
            new Dictionary<TimePeriodValues, XLTimePeriod>
            {
                { TimePeriodValues.Yesterday, XLTimePeriod.Yesterday },
                { TimePeriodValues.Today, XLTimePeriod.Today },
                { TimePeriodValues.Tomorrow, XLTimePeriod.Tomorrow },
                { TimePeriodValues.Last7Days, XLTimePeriod.InTheLast7Days },
                { TimePeriodValues.LastWeek, XLTimePeriod.LastWeek },
                { TimePeriodValues.ThisWeek, XLTimePeriod.ThisWeek },
                { TimePeriodValues.NextWeek, XLTimePeriod.NextWeek },
                { TimePeriodValues.LastMonth, XLTimePeriod.LastMonth },
                { TimePeriodValues.ThisMonth, XLTimePeriod.ThisMonth },
                { TimePeriodValues.NextMonth, XLTimePeriod.NextMonth },
            };

        public static XLTimePeriod ToClosedXml(this TimePeriodValues value)
        {
            return TimePeriodMap[value];
        }

        private static readonly IReadOnlyDictionary<PivotAreaValues, XLPivotAreaType> PivotAreaMap =
            new Dictionary<PivotAreaValues, XLPivotAreaType>
            {
                { PivotAreaValues.None, XLPivotAreaType.None },
                { PivotAreaValues.Normal, XLPivotAreaType.Normal },
                { PivotAreaValues.Data, XLPivotAreaType.Data },
                { PivotAreaValues.All, XLPivotAreaType.All },
                { PivotAreaValues.Origin, XLPivotAreaType.Origin },
                { PivotAreaValues.Button, XLPivotAreaType.Button },
                { PivotAreaValues.TopRight, XLPivotAreaType.TopRight },
                { PivotAreaValues.TopEnd, XLPivotAreaType.TopEnd },
            };

        public static XLPivotAreaType ToClosedXml(this PivotAreaValues value)
        {
            return PivotAreaMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.SparklineTypeValues, XLSparklineType> SparklineTypeMap =
            new Dictionary<X14.SparklineTypeValues, XLSparklineType>
            {
                { X14.SparklineTypeValues.Line, XLSparklineType.Line },
                { X14.SparklineTypeValues.Column, XLSparklineType.Column },
                { X14.SparklineTypeValues.Stacked, XLSparklineType.Stacked },
            };

        public static XLSparklineType ToClosedXml(this X14.SparklineTypeValues value)
        {
            return SparklineTypeMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.SparklineAxisMinMaxValues, XLSparklineAxisMinMax> SparklineAxisMinMaxMap =
            new Dictionary<X14.SparklineAxisMinMaxValues, XLSparklineAxisMinMax>
            {
                { X14.SparklineAxisMinMaxValues.Individual, XLSparklineAxisMinMax.Automatic },
                { X14.SparklineAxisMinMaxValues.Group, XLSparklineAxisMinMax.SameForAll },
                { X14.SparklineAxisMinMaxValues.Custom, XLSparklineAxisMinMax.Custom },
            };

        public static XLSparklineAxisMinMax ToClosedXml(this X14.SparklineAxisMinMaxValues value)
        {
            return SparklineAxisMinMaxMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.DisplayBlanksAsValues, XLDisplayBlanksAsValues> DisplayBlanksAsMap =
            new Dictionary<X14.DisplayBlanksAsValues, XLDisplayBlanksAsValues>
            {
                { X14.DisplayBlanksAsValues.Span, XLDisplayBlanksAsValues.Interpolate },
                { X14.DisplayBlanksAsValues.Gap, XLDisplayBlanksAsValues.NotPlotted },
                { X14.DisplayBlanksAsValues.Zero, XLDisplayBlanksAsValues.Zero },
            };

        public static XLDisplayBlanksAsValues ToClosedXml(this X14.DisplayBlanksAsValues value)
        {
            return DisplayBlanksAsMap[value];
        }

        private static readonly IReadOnlyDictionary<FieldSortValues, XLPivotSortType> FieldSortMap =
            new Dictionary<FieldSortValues, XLPivotSortType>
            {
                { FieldSortValues.Manual, XLPivotSortType.Default },
                { FieldSortValues.Ascending, XLPivotSortType.Ascending },
                { FieldSortValues.Descending, XLPivotSortType.Descending },
            };

        public static XLPivotSortType ToClosedXml(this FieldSortValues value)
        {
            return FieldSortMap[value];
        }

        private static readonly IReadOnlyDictionary<PivotTableAxisValues, XLPivotAxis> PivotTableAxisMap =
            new Dictionary<PivotTableAxisValues, XLPivotAxis>
            {
                { PivotTableAxisValues.AxisRow, XLPivotAxis.AxisRow },
                { PivotTableAxisValues.AxisColumn, XLPivotAxis.AxisCol },
                { PivotTableAxisValues.AxisPage, XLPivotAxis.AxisPage },
                { PivotTableAxisValues.AxisValues, XLPivotAxis.AxisValues },
            };

        internal static XLPivotAxis ToClosedXml(this PivotTableAxisValues value)
        {
            return PivotTableAxisMap[value];
        }

        private static readonly IReadOnlyDictionary<ItemValues, XLPivotItemType> ItemMap =
            new Dictionary<ItemValues, XLPivotItemType>
            {
                { ItemValues.Data, XLPivotItemType.Data },
                { ItemValues.Default, XLPivotItemType.Default },
                { ItemValues.Sum, XLPivotItemType.Sum },
                { ItemValues.CountA, XLPivotItemType.CountA },
                { ItemValues.Average, XLPivotItemType.Avg },
                { ItemValues.Maximum, XLPivotItemType.Max },
                { ItemValues.Minimum, XLPivotItemType.Min },
                { ItemValues.Product, XLPivotItemType.Product },
                { ItemValues.Count, XLPivotItemType.Count },
                { ItemValues.StandardDeviation, XLPivotItemType.StdDev },
                { ItemValues.StandardDeviationP, XLPivotItemType.StdDevP },
                { ItemValues.Variance, XLPivotItemType.Var },
                { ItemValues.VarianceP, XLPivotItemType.VarP },
                { ItemValues.Grand, XLPivotItemType.Grand },
                { ItemValues.Blank, XLPivotItemType.Blank },
            };

        internal static XLPivotItemType ToClosedXml(this ItemValues value)
        {
            return ItemMap[value];
        }

        private static readonly IReadOnlyDictionary<FormatActionValues, XLPivotFormatAction> FormatActionMap =
            new Dictionary<FormatActionValues, XLPivotFormatAction>
            {
                { FormatActionValues.Blank, XLPivotFormatAction.Blank },
                { FormatActionValues.Formatting, XLPivotFormatAction.Formatting },
            };

        internal static XLPivotFormatAction ToClosedXml(this FormatActionValues value)
        {
            return FormatActionMap[value];
        }

        private static readonly IReadOnlyDictionary<ScopeValues, XLPivotCfScope> ScopeMap =
            new Dictionary<ScopeValues, XLPivotCfScope>
            {
                { ScopeValues.Selection, XLPivotCfScope.SelectedCells },
                { ScopeValues.Data, XLPivotCfScope.DataFields },
                { ScopeValues.Field, XLPivotCfScope.FieldIntersections },
            };

        internal static XLPivotCfScope ToClosedXml(this ScopeValues value)
        {
            return ScopeMap[value];
        }

        private static readonly IReadOnlyDictionary<RuleValues, XLPivotCfRuleType> RuleMap =
            new Dictionary<RuleValues, XLPivotCfRuleType>
            {
                { RuleValues.None, XLPivotCfRuleType.None },
                { RuleValues.All, XLPivotCfRuleType.All },
                { RuleValues.Row, XLPivotCfRuleType.Row },
                { RuleValues.Column, XLPivotCfRuleType.Column },
            };

        internal static XLPivotCfRuleType ToClosedXml(this RuleValues value)
        {
            return RuleMap[value];
        }

        #endregion To ClosedXml
    }
}
