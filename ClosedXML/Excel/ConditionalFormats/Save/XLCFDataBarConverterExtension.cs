using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;

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
                Id = cf.Name
            };

            DataBar dataBar = new DataBar()
            {
                MinLength = 0,
                MaxLength = 100,
                Gradient = false,
                AxisPosition = DataBarAxisPositionValues.Middle
            };

            var cfMin = new ConditionalFormattingValueObject { Type = ConditionalFormattingValueObjectTypeValues.AutoMin };
            var cfMax = new ConditionalFormattingValueObject() { Type = ConditionalFormattingValueObjectTypeValues.AutoMax };

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
    }
}
