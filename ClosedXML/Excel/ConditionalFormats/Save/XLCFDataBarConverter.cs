using ClosedXML.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            var dataBar = new DataBar { ShowValue = !cf.ShowBarOnly };

            var conditionalFormatValueObject1 = new ConditionalFormatValueObject { Type = cf.ContentTypes[1].ToOpenXml() };
            if (cf.Values.Any() && cf.Values[1]?.Value != null) conditionalFormatValueObject1.Val = cf.Values[1].Value;

            var conditionalFormatValueObject2 = new ConditionalFormatValueObject { Type = cf.ContentTypes[2].ToOpenXml() };
            if (cf.Values.Count >= 2 && cf.Values[2]?.Value != null) conditionalFormatValueObject2.Val = cf.Values[2].Value;

            var color = new Color();
            switch (cf.Colors[1].ColorType)
            {
                case XLColorType.Color:
                    color.Rgb = cf.Colors[1].Color.ToHex();
                    break;

                case XLColorType.Theme:
                    color.Theme = System.Convert.ToUInt32(cf.Colors[1].ThemeColor);
                    break;

                case XLColorType.Indexed:
                    color.Indexed = System.Convert.ToUInt32(cf.Colors[1].Indexed);
                    break;
            }

            dataBar.Append(conditionalFormatValueObject1);
            dataBar.Append(conditionalFormatValueObject2);
            dataBar.Append(color);

            conditionalFormattingRule.Append(dataBar);

            if (cf.Colors.Count > 1)
            {
                ConditionalFormattingRuleExtensionList conditionalFormattingRuleExtensionList = new ConditionalFormattingRuleExtensionList();

                ConditionalFormattingRuleExtension conditionalFormattingRuleExtension = new ConditionalFormattingRuleExtension { Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" };
                conditionalFormattingRuleExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                DocumentFormat.OpenXml.Office2010.Excel.Id id = new DocumentFormat.OpenXml.Office2010.Excel.Id
                {
                    Text = (cf as XLConditionalFormat).Id.WrapInBraces()
                };
                conditionalFormattingRuleExtension.Append(id);

                conditionalFormattingRuleExtensionList.Append(conditionalFormattingRuleExtension);

                conditionalFormattingRule.Append(conditionalFormattingRuleExtensionList);
            }
            return conditionalFormattingRule;
        }
    }
}
