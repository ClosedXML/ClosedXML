using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using System;

namespace ClosedXML.Excel.IO
{
    internal class CustomFilePropertiesPartWriter
    {
        internal static void GenerateContent(CustomFilePropertiesPart customFilePropertiesPart, XLWorkbook workbook)
        {
            var properties = new Properties();
            properties.AddNamespaceDeclaration("vt",
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            var propertyId = 1;
            foreach (var p in workbook.CustomProperties)
            {
                propertyId++;
                var customDocumentProperty = new CustomDocumentProperty
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    PropertyId = propertyId,
                    Name = p.Name
                };
                if (p.Type == XLCustomPropertyType.Text)
                {
                    var vTlpwstr1 = new VTLPWSTR { Text = p.GetValue<string>() };
                    customDocumentProperty.AppendChild(vTlpwstr1);
                }
                else if (p.Type == XLCustomPropertyType.Date)
                {
                    var vTFileTime1 = new VTFileTime
                    {
                        Text =
                            p.GetValue<DateTime>().ToUniversalTime().ToString(
                                "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")
                    };
                    customDocumentProperty.AppendChild(vTFileTime1);
                }
                else if (p.Type == XLCustomPropertyType.Number)
                {
                    var vTDouble1 = new VTDouble
                    {
                        Text = p.GetValue<Double>().ToInvariantString()
                    };
                    customDocumentProperty.AppendChild(vTDouble1);
                }
                else
                {
                    var vTBool1 = new VTBool { Text = p.GetValue<Boolean>().ToString().ToLower() };
                    customDocumentProperty.AppendChild(vTBool1);
                }
                properties.AppendChild(customDocumentProperty);
            }

            customFilePropertiesPart.Properties = properties;
        }
    }
}
