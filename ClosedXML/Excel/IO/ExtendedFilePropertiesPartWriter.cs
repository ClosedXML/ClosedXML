#nullable disable

using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.IO
{
    internal class ExtendedFilePropertiesPartWriter
    {
        internal static void GenerateContent(ExtendedFilePropertiesPart extendedFilePropertiesPart, XLWorkbook workbook)
        {
            if (extendedFilePropertiesPart.Properties == null)
                extendedFilePropertiesPart.Properties = new Properties();

            var properties = extendedFilePropertiesPart.Properties;
            if (
                !properties.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("vt",
                    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")))
            {
                properties.AddNamespaceDeclaration("vt",
                    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            }

            if (properties.Application == null)
                properties.AppendChild(new Application { Text = "Microsoft Excel" });

            if (properties.DocumentSecurity == null)
                properties.AppendChild(new DocumentSecurity { Text = "0" });

            if (properties.ScaleCrop == null)
                properties.AppendChild(new ScaleCrop { Text = "false" });

            if (properties.HeadingPairs == null)
                properties.HeadingPairs = new HeadingPairs();

            if (properties.TitlesOfParts == null)
                properties.TitlesOfParts = new TitlesOfParts();

            properties.HeadingPairs.VTVector = new VTVector { BaseType = VectorBaseValues.Variant };

            properties.TitlesOfParts.VTVector = new VTVector { BaseType = VectorBaseValues.Lpstr };

            var vTVectorOne = properties.HeadingPairs.VTVector;

            var vTVectorTwo = properties.TitlesOfParts.VTVector;

            var modifiedWorksheets =
                ((IEnumerable<XLWorksheet>)workbook.WorksheetsInternal).Select(w => new { w.Name, Order = w.Position }).ToList();
            var modifiedNamedRanges = GetModifiedNamedRanges(workbook);
            var modifiedWorksheetsCount = modifiedWorksheets.Count;
            var modifiedNamedRangesCount = modifiedNamedRanges.Count;

            InsertOnVtVector(vTVectorOne, "Worksheets", 0, modifiedWorksheetsCount.ToInvariantString());
            InsertOnVtVector(vTVectorOne, "Named Ranges", 2, modifiedNamedRangesCount.ToInvariantString());

            vTVectorTwo.Size = (UInt32)(modifiedNamedRangesCount + modifiedWorksheetsCount);

            foreach (
                var vTlpstr3 in modifiedWorksheets.OrderBy(w => w.Order).Select(w => new VTLPSTR { Text = w.Name }))
                vTVectorTwo.AppendChild(vTlpstr3);

            foreach (var vTlpstr7 in modifiedNamedRanges.Select(nr => new VTLPSTR { Text = nr }))
                vTVectorTwo.AppendChild(vTlpstr7);

            if (workbook.Properties.Manager != null)
            {
                if (!String.IsNullOrWhiteSpace(workbook.Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Manager();

                    properties.Manager.Text = workbook.Properties.Manager;
                }
                else
                    properties.Manager = null;
            }

            if (workbook.Properties.Company == null) return;

            if (!String.IsNullOrWhiteSpace(workbook.Properties.Company))
            {
                if (properties.Company == null)
                    properties.Company = new Company();

                properties.Company.Text = workbook.Properties.Company;
            }
            else
                properties.Company = null;
        }

        private static void InsertOnVtVector(VTVector vTVector, String property, Int32 index, String text)
        {
            var m = from e1 in vTVector.Elements<Variant>()
                    where e1.Elements<VTLPSTR>().Any(e2 => e2.Text == property)
                    select e1;
            if (!m.Any())
            {
                if (vTVector.Size == null)
                    vTVector.Size = new UInt32Value(0U);

                vTVector.Size += 2U;
                var variant1 = new Variant();
                var vTlpstr1 = new VTLPSTR { Text = property };
                variant1.AppendChild(vTlpstr1);
                vTVector.InsertAt(variant1, index);

                var variant2 = new Variant();
                var vTInt321 = new VTInt32();
                variant2.AppendChild(vTInt321);
                vTVector.InsertAt(variant2, index + 1);
            }

            var targetIndex = 0;
            foreach (var e in vTVector.Elements<Variant>())
            {
                if (e.Elements<VTLPSTR>().Any(e2 => e2.Text == property))
                {
                    vTVector.ElementAt(targetIndex + 1).GetFirstChild<VTInt32>().Text = text;
                    break;
                }
                targetIndex++;
            }
        }

        private static List<string> GetModifiedNamedRanges(XLWorkbook workbook)
        {
            var namedRanges = new List<String>();
            foreach (var sheet in workbook.WorksheetsInternal)
            {
                namedRanges.AddRange(sheet.DefinedNames.Select<XLDefinedName, string>(n => sheet.Name + "!" + n.Name));
                namedRanges.Add(sheet.Name + "!Print_Area");
                namedRanges.Add(sheet.Name + "!Print_Titles");
            }
            namedRanges.AddRange(workbook.DefinedNamesInternal.Select<XLDefinedName, string>(n => n.Name));
            return namedRanges;
        }
    }
}
