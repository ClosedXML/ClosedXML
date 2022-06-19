using ClosedXML.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            var dataBar = new DataBar { ShowValue = !cf.ShowBarOnly };

            var conditionalFormatValueObject1 = GetConditionalFormatValueObjectByIndex(cf, 1, ConditionalFormatValueObjectValues.Min);
            var conditionalFormatValueObject2 = GetConditionalFormatValueObjectByIndex(cf, 2, ConditionalFormatValueObjectValues.Max);
            
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

            var conditionalFormattingRuleExtensionList = new ConditionalFormattingRuleExtensionList();
            conditionalFormattingRuleExtensionList.Append(BuildRuleExtension(cf));
            conditionalFormattingRule.Append(conditionalFormattingRuleExtensionList);

            return conditionalFormattingRule;
        }

        private ConditionalFormattingRuleExtension BuildRuleExtension(IXLConditionalFormat cf)
        {
            var conditionalFormattingRuleExtension = new ConditionalFormattingRuleExtension { Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" };
            conditionalFormattingRuleExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var id = new DocumentFormat.OpenXml.Office2010.Excel.Id
            {
                Text = (cf as XLConditionalFormat).Id.WrapInBraces()
            };
            conditionalFormattingRuleExtension.Append(id);

            return conditionalFormattingRuleExtension;
        }

        private ConditionalFormatValueObject GetConditionalFormatValueObjectByIndex(IXLConditionalFormat cf, int index, ConditionalFormatValueObjectValues defaultType)
        {
            var conditionalFormatValueObject = new ConditionalFormatValueObject();

            if (cf.ContentTypes.TryGetValue(index, out var contentType))
            {
                conditionalFormatValueObject.Type = contentType.ToOpenXml();
            }
            else
            {
                conditionalFormatValueObject.Type = defaultType;
            }

            if (cf.Values.TryGetValue(index, out var value1) && value1?.Value != null)
            {
                conditionalFormatValueObject.Val = value1.Value;
            }

            return conditionalFormatValueObject;
        }
    }
}
