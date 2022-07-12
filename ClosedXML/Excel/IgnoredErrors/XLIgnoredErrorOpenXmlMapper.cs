using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal static class XLIgnoredErrorOpenXmlMapper
    {
        internal static XLIgnoredErrorType? GetIgnoredErrorType(IgnoredError ignoredError)
        {
            if (ignoredError.CalculatedColumn ?? false == true)
                return XLIgnoredErrorType.CalculatedColumn;
            if (ignoredError.EmptyCellReference ?? false == true)
                return XLIgnoredErrorType.EmptyCellReference;
            if (ignoredError.EvalError ?? false == true)
                return XLIgnoredErrorType.EvalError;
            if (ignoredError.Formula ?? false == true)
                return XLIgnoredErrorType.Formula;
            if (ignoredError.FormulaRange ?? false == true)
                return XLIgnoredErrorType.FormulaRange;
            if (ignoredError.ListDataValidation ?? false == true)
                return XLIgnoredErrorType.ListDataValidation;
            if (ignoredError.NumberStoredAsText ?? false == true)
                return XLIgnoredErrorType.NumberAsText;
            if (ignoredError.TwoDigitTextYear ?? false == true)
                return XLIgnoredErrorType.TwoDigitTextYear;
            if (ignoredError.UnlockedFormula ?? false == true)
                return XLIgnoredErrorType.UnlockedFormula;

            return null;
        }

        internal static bool IsKnownIgnoredError(IgnoredError ignoredError)
        {
            return GetIgnoredErrorType(ignoredError) != null;
        }

        internal static void AddIgnoredErrorFromOpenXml(XLWorksheet ws, IgnoredError ignoredError)
        {
            var type = GetIgnoredErrorType(ignoredError);
            if (type == null)
                return;

            var ranges = ws.Ranges(ignoredError.SequenceOfReferences.InnerText.Replace(" ", ","));
            foreach (var range in ranges)
                ws.IgnoredErrors.Add(type.Value, range);
        }

        internal static IEnumerable<IgnoredError> GetOpenXmlIgnoredErrors(IXLWorksheet ws)
        {
            var ignoredErrors = new List<IgnoredError>();

            if (!ws.IgnoredErrors.Any())
                return ignoredErrors;

            foreach (var item in ws.IgnoredErrors.GroupBy(x => x.Type))
            {
                var seqRef = string.Join(" ", item.Select(x => x.Range.RangeAddress.FirstAddress.Equals(x.Range.RangeAddress.LastAddress) ?
                                                                    x.Range.RangeAddress.FirstAddress.ToStringRelative(false) :
                                                                    x.Range.RangeAddress.ToStringRelative(false))
                                                  .Distinct());

                switch (item.Key)
                {
                    case XLIgnoredErrorType.CalculatedColumn:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            CalculatedColumn = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.EmptyCellReference:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            EmptyCellReference = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.EvalError:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            EvalError = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.Formula:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            Formula = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.FormulaRange:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            FormulaRange = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.ListDataValidation:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            ListDataValidation = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.NumberAsText:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            NumberStoredAsText = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.TwoDigitTextYear:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            TwoDigitTextYear = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    case XLIgnoredErrorType.UnlockedFormula:
                        ignoredErrors.Add(new IgnoredError()
                        {
                            UnlockedFormula = true,
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = seqRef
                            }
                        });
                        break;
                    default:
                        throw new ArgumentOutOfRangeException($"XLIgnoredErrorType: {item.Key}");
                };
            }

            return ignoredErrors;
        }
    }
}
