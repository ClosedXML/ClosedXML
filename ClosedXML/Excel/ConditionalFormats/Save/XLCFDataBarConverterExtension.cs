using ClosedXML.Extensions;
using DocumentFormat.OpenXml.Office.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarConverterExtension : IXLCFConverterExtension
    {
        public XLCFDataBarConverterExtension()
        {
        }

        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context)
        {
            ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule()
            {
                Type = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValues.DataBar,
                Id = (cf as XLConditionalFormat).Id.WrapInBraces()
            };

            DataBar dataBar = new DataBar()
            {
                MinLength = 0,
                MaxLength = 100,
                Gradient = false,
                AxisPosition = DataBarAxisPositionValues.Middle,
                ShowValue = !cf.ShowBarOnly
            };

            ConditionalFormattingValueObjectTypeValues cfMinType = Convert(cf.ContentTypes[1].ToOpenXml());
            var cfMin = new ConditionalFormattingValueObject { Type = cfMinType };
            if (cf.Values.Any() && cf.Values[1]?.Value != null)
            {
                cfMin.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMin.Append(new Formula() { Text = cf.Values[1].Value });
            }

            ConditionalFormattingValueObjectTypeValues cfMaxType = Convert(cf.ContentTypes[2].ToOpenXml());
            var cfMax = new ConditionalFormattingValueObject { Type = cfMaxType };
            if (cf.Values.Count >= 2 && cf.Values[2]?.Value != null)
            {
                cfMax.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMax.Append(new Formula() { Text = cf.Values[2].Value });
            }

            var barAxisColor = new BarAxisColor { Rgb = XLColor.Black.Color.ToHex() };

            var negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[1].Color.ToHex() };
            if (cf.Colors.Count == 2)
            {
                negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[2].Color.ToHex() };
            }

            dataBar.Append(cfMin);
            dataBar.Append(cfMax);

            dataBar.Append(negativeFillColor);
            dataBar.Append(barAxisColor);

            conditionalFormattingRule.Append(dataBar);

            return conditionalFormattingRule;
        }

        private ConditionalFormattingValueObjectTypeValues Convert(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues obj)
        {
            switch (obj)
            {
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Max:
                    return ConditionalFormattingValueObjectTypeValues.AutoMax;
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Min:
                    return ConditionalFormattingValueObjectTypeValues.AutoMin;
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Number:
                    return ConditionalFormattingValueObjectTypeValues.Numeric;
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percent:
                    return ConditionalFormattingValueObjectTypeValues.Percent;
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percentile:
                    return ConditionalFormattingValueObjectTypeValues.Percentile;
                case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Formula:
                    return ConditionalFormattingValueObjectTypeValues.Formula;
                default:
                    throw new NotImplementedException();
            }
        }
    }
}
