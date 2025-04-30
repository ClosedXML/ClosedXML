using ClosedXML.Extensions;
using ClosedXML.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ClosedXML.Excel.IO;

internal class PivotTableCacheRecordsPartReader
{
    internal static void ReadRecords(PivotCacheRecords recordsPart, XLPivotCache pivotCache)
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
                            throw PartStructureException.InvalidAttributeFormat();

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
                            throw PartStructureException.InvalidAttributeValue();

                        fieldValues.AddIndex(index);
                        break;

                    default:
                        throw PartStructureException.ExpectedElementNotFound();
                }
            }
        }
    }
}
