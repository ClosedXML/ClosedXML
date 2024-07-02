#nullable disable

using System;
using System.Linq;
using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.IO
{
    internal class PivotTableCacheDefinitionPartReader
    {
        internal static void Load(WorkbookPart workbookPart, XLWorkbook workbook)
        {
            foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                if (pivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.WorksheetSource != null)
                {
                    var pivotSourceReference = ParsePivotSourceReference(pivotTableCacheDefinitionPart);
                    if (pivotSourceReference == null)
                        // We don't support external sources
                        continue;

                    var pivotCache = workbook.PivotCachesInternal.Add(pivotSourceReference);

                    // If WorkbookCacheRelId already has a value, it means the pivot source is being reused
                    if (string.IsNullOrWhiteSpace(pivotCache.WorkbookCacheRelId))
                    {
                        pivotCache.WorkbookCacheRelId = workbookPart.GetIdOfPart(pivotTableCacheDefinitionPart);
                    }

                    var cacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;
                    if (cacheDefinition.MissingItemsLimit is not null)
                    {
                        if (cacheDefinition.MissingItemsLimit == 0U)
                        {
                            pivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
                        }
                        else if (cacheDefinition.MissingItemsLimit == XLHelper.MaxRowNumber)
                        {
                            pivotCache.ItemsToRetainPerField = XLItemsToRetain.Max;
                        }
                    }

                    if (pivotTableCacheDefinitionPart.PivotCacheDefinition?.CacheFields is { } cacheFields)
                    {
                        ReadCacheFields(cacheFields, pivotCache);
                        if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } recordsPart)
                        {
                            ReadRecords(recordsPart, pivotCache);
                        }
                    }

                    if (pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData != null)
                    {
                        pivotCache.SaveSourceData = pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData.Value;
                    }
                }
            }
        }

        internal static XLPivotSourceReference ParsePivotSourceReference(PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
        {
            // TODO: Implement other sources besides worksheetSource
            // But for now assume names and references point directly to a range
            var wss = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheSource.WorksheetSource;

            if (!String.IsNullOrEmpty(wss.Id))
            {
                var externalRelationship = pivotTableCacheDefinitionPart.ExternalRelationships.FirstOrDefault(er => er.Id.Equals(wss.Id));
                if (externalRelationship?.IsExternal ?? false)
                {
                    // We don't support external sources
                    return null;
                }
            }

            // Source data of pivot cache are from a table or a named range.
            if (wss.Name is not null)
            {
                return new XLPivotSourceReference(wss.Name);
            }

            // Source data of pivot cache are from an area of a workbook.
            if (wss.Reference is not null && wss.Sheet is not null)
            {
                var bookArea = new XLBookArea(wss.Sheet, XLSheetRange.Parse(wss.Reference));
                return new XLPivotSourceReference(bookArea);
            }

            throw PartStructureException.MissingAttribute();
        }

        private static void ReadCacheFields(CacheFields cacheFields, XLPivotCache pivotCache)
        {
            foreach (var cacheField in cacheFields.Elements<CacheField>())
            {
                if (cacheField.Name?.Value is not { } fieldName)
                    throw PartStructureException.MissingAttribute();

                if (pivotCache.ContainsField(fieldName))
                {
                    // We don't allow duplicate field names... but what do we do if we find one? Let's just skip it.
                    continue;
                }

                var fieldStats = ReadCacheFieldStats(cacheField);
                var fieldSharedItems = cacheField.SharedItems is not null
                    ? ReadSharedItems(cacheField)
                    : new XLPivotCacheSharedItems();

                var fieldValues = new XLPivotCacheValues(fieldSharedItems, fieldStats);
                pivotCache.AddCachedField(fieldName, fieldValues);
            }
        }

        private static XLPivotCacheValuesStats ReadCacheFieldStats(CacheField cacheField)
        {
            var sharedItems = cacheField.SharedItems;

            // Various statistics about the records of the field, not just shared items.
            var containsBlank = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsBlank, false);
            var containsNumber = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsNumber, false);
            var containsOnlyInteger = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsInteger, false);
            var minValue = sharedItems?.MinValue?.Value;
            var maxValue = sharedItems?.MaxValue?.Value;
            var containsDate = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsDate, false);
            var minDate = sharedItems?.MinDate?.Value;
            var maxDate = sharedItems?.MaxDate?.Value;
            var containsString = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsString, true);
            var longText = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.LongText, false);

            // The containsMixedTypes, containsNonDate and containsSemiMixedTypes are derived from primary stats.
            return new XLPivotCacheValuesStats(
                containsBlank,
                containsNumber,
                containsOnlyInteger,
                minValue,
                maxValue,
                containsString,
                longText,
                containsDate,
                minDate,
                maxDate);
        }

        private static XLPivotCacheSharedItems ReadSharedItems(CacheField cacheField)
        {
            var sharedItems = new XLPivotCacheSharedItems();

            // If there are no shared items, the cache record can't contain field items
            // referencing the shared items.
            if (cacheField.SharedItems is not { } fieldSharedItems)
                return sharedItems;

            foreach (var item in fieldSharedItems.Elements())
            {
                // Shared items can't contain element of type index (`x`),
                // because index references shared items. That is main reason
                // for rather significant duplication with reading records.
                switch (item)
                {
                    case MissingItem:
                        sharedItems.AddMissing();
                        break;

                    case NumberItem numberItem:
                        if (numberItem.Val?.Value is not { } number)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddNumber(number);
                        break;

                    case BooleanItem booleanItem:
                        if (booleanItem.Val?.Value is not { } boolean)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddBoolean(boolean);
                        break;

                    case ErrorItem errorItem:
                        if (errorItem.Val?.Value is not { } errorText)
                            throw PartStructureException.MissingAttribute();

                        if (!XLErrorParser.TryParseError(errorText, out var error))
                            throw PartStructureException.IncorrectAttributeFormat();

                        sharedItems.AddError(error);
                        break;

                    case StringItem stringItem:
                        if (stringItem.Val?.Value is not { } text)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddString(text);
                        break;

                    case DateTimeItem dateTimeItem:
                        if (dateTimeItem.Val?.Value is not { } dateTime)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddDateTime(dateTime);
                        break;

                    default:
                        throw PartStructureException.ExpectedElementNotFound();
                }
            }

            return sharedItems;
        }

        private static void ReadRecords(PivotCacheRecords recordsPart, XLPivotCache pivotCache)
        {
            // Number of records can be rather large, preallocate capacity to avoid reallocation.
            var recordCount = recordsPart.Count?.Value is not null
                ? checked((int)recordsPart.Count.Value)
                : 0;
            pivotCache.AllocateRecordCapacity(recordCount);

            var fieldsCount = pivotCache.FieldCount;
            foreach (var record in recordsPart.Elements<PivotCacheRecord>())
            {
                var recordColumns = record.ChildElements.Count;
                if (recordColumns != fieldsCount)
                    throw PartStructureException.IncorrectElementsCount();

                for (var fieldIdx = 0; fieldIdx < fieldsCount; ++fieldIdx)
                {
                    var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                    var recordItem = record.ElementAt(fieldIdx);

                    // Don't add values to the shared items of a cache when record value is added, because we want 1:1
                    // read/write. Read them from definition. Whatever is in shared items now should be written out,
                    // unless there is a cache refresh. Basically trust the author of the workbook that it is valid.
                    switch (recordItem)
                    {
                        case MissingItem:
                            fieldValues.AddMissing();
                            break;

                        case NumberItem numberItem:
                            if (numberItem.Val?.Value is not { } number)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddNumber(number);
                            break;

                        case BooleanItem booleanItem:
                            if (booleanItem.Val?.Value is not { } boolean)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddBoolean(boolean);
                            break;

                        case ErrorItem errorItem:
                            if (errorItem.Val?.Value is not { } errorText)
                                throw PartStructureException.MissingAttribute();

                            if (!XLErrorParser.TryParseError(errorText, out var error))
                                throw PartStructureException.IncorrectAttributeFormat();

                            fieldValues.AddError(error);
                            break;

                        case StringItem stringItem:
                            if (stringItem.Val?.Value is not { } text)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddString(text);
                            break;

                        case DateTimeItem dateTimeItem:
                            if (dateTimeItem.Val?.Value is not { } dateTime)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddDateTime(dateTime);
                            break;

                        case FieldItem indexItem:
                            if (indexItem.Val?.Value is not { } index)
                                throw PartStructureException.MissingAttribute();

                            if (index >= fieldValues.SharedCount)
                                throw PartStructureException.IncorrectAttributeValue();

                            fieldValues.AddIndex(index);
                            break;

                        default:
                            throw PartStructureException.ExpectedElementNotFound();
                    }
                }
            }
        }
    }
}
