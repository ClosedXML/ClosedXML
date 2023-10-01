#nullable disable

using ClosedXML.Excel.Cells;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Extensions;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class PivotTableCacheDefinitionPartWriter
    {
        internal static void GenerateContent(
            WorkbookPart workbookPart,
            XLPivotCache pivotCache,
            SaveContext context)
        {
            Debug.Assert(workbookPart.Workbook.PivotCaches is not null);
            Debug.Assert(!string.IsNullOrEmpty(pivotCache.WorkbookCacheRelId));

            var pivotTableCacheDefinitionPart = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(pivotCache.WorkbookCacheRelId);

            var pivotCacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;

            if (pivotCacheDefinition == null)
            {
                pivotCacheDefinition = new PivotCacheDefinition { Id = "rId1" };

                pivotCacheDefinition.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                pivotTableCacheDefinitionPart.PivotCacheDefinition = pivotCacheDefinition;
            }

            #region CreatedVersion

            byte createdVersion = XLConstants.PivotTable.CreatedVersion;

            if (pivotCacheDefinition.CreatedVersion?.HasValue ?? false)
                pivotCacheDefinition.CreatedVersion = Math.Max(createdVersion, pivotCacheDefinition.CreatedVersion.Value);
            else
                pivotCacheDefinition.CreatedVersion = createdVersion;

            #endregion CreatedVersion

            #region RefreshedVersion

            byte refreshedVersion = XLConstants.PivotTable.RefreshedVersion;
            if (pivotCacheDefinition.RefreshedVersion?.HasValue ?? false)
                pivotCacheDefinition.RefreshedVersion = Math.Max(refreshedVersion, pivotCacheDefinition.RefreshedVersion.Value);
            else
                pivotCacheDefinition.RefreshedVersion = refreshedVersion;

            #endregion RefreshedVersion

            #region MinRefreshableVersion

            byte minRefreshableVersion = 3;
            if (pivotCacheDefinition.MinRefreshableVersion?.HasValue ?? false)
                pivotCacheDefinition.MinRefreshableVersion = Math.Max(minRefreshableVersion, pivotCacheDefinition.MinRefreshableVersion.Value);
            else
                pivotCacheDefinition.MinRefreshableVersion = minRefreshableVersion;

            #endregion MinRefreshableVersion

            pivotCacheDefinition.SaveData = pivotCache.SaveSourceData;
            pivotCacheDefinition.RefreshOnLoad = true; //pt.RefreshDataOnOpen

            var pivotSourceInfo = new PivotSourceInfo
            {
                Guid = pivotCache.Guid,
                Fields = new Dictionary<String, PivotTableFieldInfo>()
            };

            if (pivotCache.ItemsToRetainPerField == XLItemsToRetain.None)
                pivotCacheDefinition.MissingItemsLimit = 0U;
            else if (pivotCache.ItemsToRetainPerField == XLItemsToRetain.Max)
                pivotCacheDefinition.MissingItemsLimit = XLHelper.MaxRowNumber;

            // Begin CacheSource
            var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
            var worksheetSource = new WorksheetSource();

            switch (pivotCache.PivotSourceReference.SourceType)
            {
                case XLPivotTableSourceType.Range:
                    worksheetSource.Name = null;
                    worksheetSource.Reference = pivotCache.PivotSourceReference.SourceRange.RangeAddress.ToStringRelative(includeSheet: false);

                    // Do not quote worksheet name with whitespace here - issue #955
                    worksheetSource.Sheet = pivotCache.PivotSourceReference.SourceRange.RangeAddress.Worksheet.Name;
                    break;

                case XLPivotTableSourceType.Table:
                    worksheetSource.Name = pivotCache.PivotSourceReference.SourceTable.Name;
                    worksheetSource.Reference = null;
                    worksheetSource.Sheet = null;
                    break;

                default:
                    throw new NotSupportedException($"Pivot table source type {pivotCache.PivotSourceReference.SourceType} is not supported.");
            }

            cacheSource.AppendChild(worksheetSource);
            pivotCacheDefinition.CacheSource = cacheSource;

            // End CacheSource

            // Begin CacheFields
            var cacheFields = pivotCacheDefinition.CacheFields;
            if (cacheFields == null)
            {
                cacheFields = new CacheFields();
                pivotCacheDefinition.CacheFields = cacheFields;
            }

            foreach (var cacheFieldName in pivotCache.FieldNames)
            {
                var fieldValues = pivotCache.GetFieldSharedItems(cacheFieldName);

                var distinctFieldValues = fieldValues
                    .GetCellValues()
                    .Distinct(XLCellValueComparer.OrdinalIgnoreCase)
                    .ToArray();

                var types = distinctFieldValues
                    .Select(v => v.Type)
                    .Distinct()
                    .ToArray();

                // .CacheFields is cleared when workbook is begin saved
                // So if there are any entries, it would be from previous pivot tables
                // with an identical source range.
                // When pivot sources get its refactoring, this will not be necessary
                var cacheField = pivotCacheDefinition
                    .CacheFields
                    .Elements<CacheField>()
                    .FirstOrDefault(f => f.Name == cacheFieldName);

                if (cacheField == null)
                {
                    cacheField = new CacheField
                    {
                        Name = cacheFieldName,
                        SharedItems = new SharedItems()
                    };
                    cacheFields.AppendChild(cacheField);
                }
                var sharedItems = cacheField.SharedItems;

                var ptfi = new PivotTableFieldInfo
                {
                    IsTotallyBlankField = fieldValues.Count == 0,
                    MixedDataType = types.Length > 1,
                    DistinctValues = distinctFieldValues,
                };

                if (types.Any())
                {
                    sharedItems.Count = null;

                    // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shareditems?view=openxml-2.8.1#remarks
                    // The following attributes are not required or used if there are no items in sharedItems.
                    // - containsBlank
                    // - containsSemiMixedTypes
                    // - containsMixedTypes
                    // - longText

                    // Specifies a boolean value that indicates whether this field contains a blank value.
                    sharedItems.ContainsBlank = OpenXmlHelper.GetBooleanValue(types.Contains(XLDataType.Blank), false);

                    var containsDate = types.Contains(XLDataType.DateTime) || types.Contains(XLDataType.TimeSpan);
                    sharedItems.ContainsDate = OpenXmlHelper.GetBooleanValue(containsDate, false);

                    // Remember: Blank is not a type in OOXML, but is a value

                    // ISO29500: Specifies a boolean value that indicates whether this field contains more than one data type.
                    // MS-OI29500: In Office, the containsMixedTypes attribute assumes that boolean and error shall be considered part of the string type.
                    var containsMixedTypes = types
                        .Where(t => t != XLDataType.Blank)
                        .Select(t => t is XLDataType.Boolean or XLDataType.Error ? XLDataType.Text : t)
                        .Distinct()
                        .Count() > 1;
                    sharedItems.ContainsMixedTypes = OpenXmlHelper.GetBooleanValue(containsMixedTypes, false);

                    // ISO29500: Specifies a boolean value that indicates that the field contains at least one value that is not a date.
                    var containsNonDate = types.Where(t => t != XLDataType.Blank).Any(t => t != XLDataType.DateTime && t != XLDataType.TimeSpan);
                    sharedItems.ContainsNonDate = OpenXmlHelper.GetBooleanValue(containsNonDate, true);

                    // Excel will have to repair the cache definition, if both @containsNumber and @containsDate are specified. Likely because
                    // ultimately they are both numbers, but date has preference.
                    var containsNumber = !containsDate && types.Contains(XLDataType.Number);
                    sharedItems.ContainsNumber = OpenXmlHelper.GetBooleanValue(containsNumber, false);

                    // @containsInteger has a prerequisite @containsNumber.
                    // MS-OI29500: In Office, @containsNumber shall be 1 or true when @containsInteger is specified. 
                    if (containsNumber)
                    {
                        // MS-OI29500: In Office, a value of 1 or true for the containsInteger attribute indicates this field contains only integer values and does not contain non - integer numeric values.
                        var onlyIntegers = ptfi.DistinctValues
                            .Where(v => v.Type == XLDataType.Number)
                            .Select(v => v.GetNumber())
                            .All(dbl => (dbl % 1) < double.Epsilon);
                        sharedItems.ContainsInteger = OpenXmlHelper.GetBooleanValue(onlyIntegers, false);
                    }

                    // ISO29500: A value of 1 or true indicates at least one text value, and can also contain a mix of other data types and blank values.
                    // MS-OI29500: Office expects that the containsSemiMixedTypes attribute is true when the field contains text, blank, boolean or error values.
                    var containsSemiMixedTypes = types.Any(t => t is XLDataType.Text or XLDataType.Blank or XLDataType.Boolean or XLDataType.Error);
                    sharedItems.ContainsSemiMixedTypes = OpenXmlHelper.GetBooleanValue(containsSemiMixedTypes, true);

                    // MS-OI29500: In Office, boolean and error are considered strings in the context of the containsString attribute.
                    var containsString = types.Any(t => t is XLDataType.Text or XLDataType.Boolean or XLDataType.Error);
                    sharedItems.ContainsString = OpenXmlHelper.GetBooleanValue(containsString, true);

                    sharedItems.Count = (UInt32)distinctFieldValues.Length;

                    var longText = types.Contains(XLDataType.Text) && ptfi.DistinctValues.Any(v => v.IsText && v.GetText().Length > 255);
                    sharedItems.LongText = OpenXmlHelper.GetBooleanValue(longText, false);

                    // @minDate/@maxDate can be present, only if at least one child is a d element.
                    if (types.Any(v => v == XLDataType.DateTime || v == XLDataType.TimeSpan))
                    {
                        // This is an exception to the "1900 is a leap year". Values are saved correctly, i.e starting at 1899-12-30. TimeSpan as well.
                        sharedItems.MinDate = DateTime.FromOADate(ptfi.DistinctValues.Where(x => x.IsUnifiedNumber).Min(v => v.GetUnifiedNumber()));
                        sharedItems.MaxDate = DateTime.FromOADate(ptfi.DistinctValues.Where(x => x.IsUnifiedNumber).Max(v => v.GetUnifiedNumber()));
                    }
                    else if (types.Contains(XLDataType.Number))
                    {
                        // MS-OI29500: Use else branch, @minValue/@maxValue shouldn't be present, if there is a @minDate/@maxDate.

                        // If the field contains a date, the number values are considered serial date times.
                        // Don't indicate that date field with numbers contains numbers, Excel would refuse to load the file
                        sharedItems.MinValue = ptfi.DistinctValues.Where(x => x.IsNumber).Min(v => v.GetNumber());
                        sharedItems.MaxValue = ptfi.DistinctValues.Where(x => x.IsNumber).Max(v => v.GetNumber());
                    }

                    foreach (var value in distinctFieldValues)
                    {
                        OpenXmlElement toAdd = value.Type switch
                        {
                            XLDataType.Blank => new MissingItem(),
                            XLDataType.Boolean => new BooleanItem { Val = value.GetBoolean() },
                            XLDataType.Number => new NumberItem { Val = value.GetNumber() },
                            XLDataType.Text => new StringItem { Val = value.GetText() },
                            XLDataType.Error => new ErrorItem { Val = value.GetError().ToDisplayString() },
                            XLDataType.DateTime => new DateTimeItem { Val = DateTime.FromOADate(value.GetUnifiedNumber()) },
                            XLDataType.TimeSpan => new DateTimeItem { Val = DateTime.FromOADate(value.GetUnifiedNumber()) },
                            _ => throw new InvalidOperationException()
                        };
                        sharedItems.AppendChild(toAdd);
                    }
                }

                pivotSourceInfo.Fields.Add(cacheFieldName, ptfi);
            }

            // End CacheFields

            var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any() ?
                pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().First() :
                pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            var pivotCacheRecords = new PivotCacheRecords();
            pivotCacheRecords.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotTableCacheRecordsPart.PivotCacheRecords = pivotCacheRecords;

            context.PivotSources.Add(pivotSourceInfo.Guid, pivotSourceInfo);
        }
    }
}
