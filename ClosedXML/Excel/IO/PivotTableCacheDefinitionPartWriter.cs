#nullable disable

using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Extensions;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class PivotTableCacheDefinitionPartWriter
    {
        internal static void GenerateContent(
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart,
            XLPivotCache pivotCache,
            SaveContext context)
        {
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
                case XLPivotTableSourceType.Area:
                    var bookArea = pivotCache.PivotSourceReference.Area!.Value;
                    worksheetSource.Name = null;
                    worksheetSource.Reference = bookArea.Area.ToString();

                    // Do not quote worksheet name with whitespace here - issue #955
                    worksheetSource.Sheet = bookArea.Name;
                    break;

                case XLPivotTableSourceType.Named:
                    worksheetSource.Name = pivotCache.PivotSourceReference.Name!;
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

            for (var fieldIdx = 0; fieldIdx < pivotCache.FieldCount; ++fieldIdx)
            {
                var cacheFieldName = pivotCache.FieldNames[fieldIdx];
                var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                var xlSharedItems = pivotCache.GetFieldSharedItems(fieldIdx)
                    .GetCellValues()
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
                    IsTotallyBlankField = xlSharedItems.Length == 0,
                    MixedDataType = xlSharedItems
                        .Select(v => v.Type)
                        .Distinct()
                        .Count() > 1,
                    DistinctValues = xlSharedItems,
                };

                var stats = fieldValues.Stats;

                sharedItems.Count = fieldValues.SharedCount != 0 ? checked((uint)xlSharedItems.Length) : null;

                // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shareditems?view=openxml-2.8.1#remarks
                // The following attributes are not required or used if there are no items in sharedItems.
                // - containsBlank
                // - containsSemiMixedTypes
                // - containsMixedTypes
                // - longText

                // Specifies a boolean value that indicates whether this field contains a blank value.
                sharedItems.ContainsBlank = OpenXmlHelper.GetBooleanValue(stats.ContainsBlank, false);

                sharedItems.ContainsDate = OpenXmlHelper.GetBooleanValue(stats.ContainsDate, false);

                // Remember: Blank is not a type in OOXML, but is a value
                var typesCount = 0;
                if (stats.ContainsNumber)
                    typesCount++;

                if (stats.ContainsString)
                    typesCount++;

                if (stats.ContainsDate)
                    typesCount++;

                // ISO29500: Specifies a boolean value that indicates whether this field contains more than one data type.
                // MS-OI29500: In Office, the containsMixedTypes attribute assumes that boolean and error shall be considered part of the string type.
                sharedItems.ContainsMixedTypes = OpenXmlHelper.GetBooleanValue(typesCount > 1, false);

                // ISO29500: Specifies a boolean value that indicates that the field contains at least one value that is not a date.
                var containsNonDate = stats.ContainsString || stats.ContainsNumber;
                sharedItems.ContainsNonDate = OpenXmlHelper.GetBooleanValue(containsNonDate, true);

                // Excel will have to repair the cache definition, if both @containsNumber and @containsDate are specified. Likely because
                // ultimately they are both numbers, but date has preference.
                if (stats.ContainsDate)
                {
                    // If the field contains a date, the number values are considered serial date times.

                    // This is an exception to the "1900 is a leap year". Values are saved correctly, i.e starting at 1899-12-30.
                    long? minValueAsDateTime = stats.MinValue is not null ? DateTime.FromOADate(stats.MinValue.Value).Ticks : null;
                    long? maxValueAsDateTime = stats.MaxValue is not null ? DateTime.FromOADate(stats.MaxValue.Value).Ticks : null;

                    long? minDateTicks = Min(stats.MinDate?.Ticks, minValueAsDateTime);
                    long? maxDateTicks = Max(stats.MaxDate?.Ticks, maxValueAsDateTime);

                    // @minDate/@maxDate can be present, only if at least one child is a d element.
                    sharedItems.MinDate = minDateTicks is not null ? new DateTime(minDateTicks.Value) : null;
                    sharedItems.MaxDate = maxDateTicks is not null ? new DateTime(maxDateTicks.Value) : null;

                    static long? Min(long? val1, long? val2)
                    {
                        if (val1 is null || val2 is null)
                            return val1 ?? val2;

                        return Math.Min(val1.Value, val2.Value);
                    }

                    static long? Max(long? val1, long? val2)
                    {
                        if (val1 is null || val2 is null)
                            return val1 ?? val2;

                        return Math.Max(val1.Value, val2.Value);
                    }
                }
                else if (stats.ContainsNumber)
                {
                    // Don't indicate that date field with numbers contains numbers, Excel would refuse to load the file
                    sharedItems.ContainsNumber = OpenXmlHelper.GetBooleanValue(stats.ContainsNumber, false);

                    // @containsInteger has a prerequisite @containsNumber, MS-OI29500: In Office, @containsNumber shall be 1 or true when @containsInteger is specified.
                    // MS-OI29500: In Office, a value of 1 or true for the containsInteger attribute indicates this field contains only integer values and does not contain non - integer numeric values.
                    sharedItems.ContainsInteger = OpenXmlHelper.GetBooleanValue(stats.ContainsInteger, false);

                    sharedItems.MinValue = stats.MinValue;
                    sharedItems.MaxValue = stats.MaxValue;
                }

                // ISO29500: A value of 1 or true indicates at least one text value, and can also contain a mix of other data types and blank values.
                // MS-OI29500: Office expects that the containsSemiMixedTypes attribute is true when the field contains text, blank, boolean or error values.
                var containsSemiMixedTypes = stats.ContainsString || stats.ContainsBlank;
                sharedItems.ContainsSemiMixedTypes = OpenXmlHelper.GetBooleanValue(containsSemiMixedTypes, true);

                // MS-OI29500: In Office, boolean and error are considered strings in the context of the containsString attribute.
                sharedItems.ContainsString = OpenXmlHelper.GetBooleanValue(stats.ContainsString, true);

                sharedItems.LongText = OpenXmlHelper.GetBooleanValue(stats.LongText, false);

                foreach (var value in xlSharedItems)
                {
                    OpenXmlElement toAdd = value.Type switch
                    {
                        XLDataType.Blank => new MissingItem(),
                        XLDataType.Boolean => new BooleanItem { Val = value.GetBoolean() },
                        XLDataType.Number => new NumberItem { Val = value.GetNumber() },
                        XLDataType.Text => new StringItem { Val = value.GetText() },
                        XLDataType.Error => new ErrorItem { Val = value.GetError().ToDisplayString() },
                        XLDataType.DateTime => new DateTimeItem { Val = value.GetDateTime() },
                        XLDataType.TimeSpan => new DateTimeItem { Val = DateTime.FromOADate(value.GetUnifiedNumber()) },
                        _ => throw new InvalidOperationException()
                    };
                    sharedItems.AppendChild(toAdd);
                }

                pivotSourceInfo.Fields.Add(cacheFieldName, ptfi);
            }

            // End CacheFields
            context.PivotSources.Add(pivotSourceInfo.Guid, pivotSourceInfo);
        }
    }
}
