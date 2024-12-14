using System;
using System.Linq;
using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
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
                var cacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;
                if (cacheDefinition.CacheSource is not { } cacheSource)
                    throw PartStructureException.RequiredElementIsMissing("cacheSource");

                var pivotSourceReference = ParsePivotSourceReference(cacheSource);
                var pivotCache = workbook.PivotCachesInternal.Add(pivotSourceReference);

                // If WorkbookCacheRelId already has a value, it means the pivot source is being reused
                if (string.IsNullOrWhiteSpace(pivotCache.WorkbookCacheRelId))
                {
                    pivotCache.WorkbookCacheRelId = workbookPart.GetIdOfPart(pivotTableCacheDefinitionPart);
                }

                if (cacheDefinition.MissingItemsLimit?.Value is { } missingItemsLimit)
                {
                    pivotCache.ItemsToRetainPerField = missingItemsLimit switch
                    {
                        0 => XLItemsToRetain.None,
                        XLHelper.MaxRowNumber => XLItemsToRetain.Max,
                        _ => XLItemsToRetain.Automatic,
                    };
                }

                if (cacheDefinition.CacheFields is { } cacheFields)
                {
                    ReadCacheFields(cacheFields, pivotCache);
                    if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } recordsPart)
                    {
                        ReadRecords(recordsPart, pivotCache);
                    }
                }

                pivotCache.SaveSourceData = cacheDefinition.SaveData?.Value ?? true;
            }
        }

        internal static IXLPivotSource ParsePivotSourceReference(CacheSource cacheSource)
        {
            // Cache source has several types. Each has a specific required format. Do not use different
            // combinations, Excel will crash or at least try to repair
            // [worksheet] uses a worksheet source:
            //   * An unnamed range in a sheet: Uses `sheet` and `ref`.
            //   * An table: Uses `name` that contains a name of the table.
            // [external]
            //   * `connectionId` link to external relationships.
            // [consolidation]
            //  * uses consolidation tag and a list of range sets plus optionally
            //    page fields to add a custom report fields that allow user to select
            //    ranges from rangeSet to calculate values.
            // [scenario]
            //  * only type attribute tag is specified, no other value. Likely linked
            //    through cacheField names (e.g. <cacheField name="$A$1 by">).

            // Not all sources are supported, but at least pipe the data through so the load/save works
            IEnumValue sourceType = cacheSource.Type?.Value ?? throw PartStructureException.MissingAttribute();
            if (sourceType.Equals(SourceValues.Worksheet))
            {
                var sheetSource = cacheSource.WorksheetSource;
                if (sheetSource is null)
                    throw PartStructureException.ExpectedElementNotFound("'worksheetSource' element is required for type 'worksheet'.");

                // If the source is a defined name, it must be a single area reference
                if (sheetSource.Name?.Value is { } tableOrName)
                {
                    if (sheetSource.Id?.Value is { } externalWorkbookRelId)
                        return new XLPivotSourceExternalWorkbook(externalWorkbookRelId, tableOrName);

                    return new XLPivotSourceReference(tableOrName);
                }

                if (sheetSource.Sheet?.Value is { } sheetName &&
                    sheetSource.Reference?.Value is { } areaRef &&
                    XLSheetRange.TryParse(areaRef.AsSpan(), out var sheetArea))
                {
                    var area = new XLBookArea(sheetName, sheetArea);
                    if (sheetSource.Id?.Value is { } externalWorkbookRelId)
                        return new XLPivotSourceExternalWorkbook(externalWorkbookRelId, area);

                    // area is in this workbook
                    return new XLPivotSourceReference(area);
                }

                throw PartStructureException.IncorrectElementFormat("worksheetSource");
            }

            if (sourceType.Equals(SourceValues.External))
            {
                if (cacheSource.ConnectionId?.Value is not { } connectionId)
                    throw PartStructureException.MissingAttribute("connectionId");

                return new XLPivotSourceConnection(connectionId);
            }

            if (sourceType.Equals(SourceValues.Consolidation))
            {
                throw new NotImplementedException();
            }

            if (sourceType.Equals(SourceValues.Scenario))
            {
                return new XLPivotSourceScenario();
            }

            throw PartStructureException.InvalidAttributeValue(sourceType.Value);
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
