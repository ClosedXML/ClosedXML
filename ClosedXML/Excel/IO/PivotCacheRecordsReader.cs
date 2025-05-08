using ClosedXML.Extensions;
using ClosedXML.IO;
using System;

namespace ClosedXML.Excel.IO;

internal partial class PivotCacheRecordsReader
{
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly XmlTreeReader _reader;
    private readonly XLPivotCache _pivotCache;

    /// <summary>
    /// Index of current field that is read from the <c>r</c> element.
    /// </summary>
    private int _fieldIdx;

    public PivotCacheRecordsReader(XmlTreeReader reader, XLPivotCache pivotCache)
    {
        _reader = reader;
        _pivotCache = pivotCache;
    }

    internal void ReadRecordsToCache()
    {
        // Don't add values to the shared items of a cache when record value is added, because we want 1:1
        // read/write. Read them from definition. Whatever is in shared items now should be written out,
        // unless there is a cache refresh. Basically trust the author of the workbook that it is valid.
        _reader.Open("pivotCacheRecords", _ns);
        var recordCount = _reader.GetCount();
        _pivotCache.AllocateRecordCapacity(recordCount);

        while (_reader.TryOpen("r", _ns))
        {
            ParseRecord("r");
        }

        if (_reader.TryOpen("extLst", _ns))
        {
            _reader.Skip("extLst");
        }

        _reader.Close("pivotCacheRecords", _ns);
    }

    partial void OnRecordParsed()
    {
        // Each record should have element for each field
        var fieldsCount = _pivotCache.FieldCount;
        if (_fieldIdx != fieldsCount)
            throw PartStructureException.IncorrectElementsCount();

        // Record was read, reset field index for next record.
        _fieldIdx = 0;
    }

    partial void OnMissingParsed(bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b)
    {
        var fieldValues = GetFieldValues();
        fieldValues.AddMissing();
    }

    partial void OnNumberParsed(double v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b)
    {
        var fieldValues = GetFieldValues();
        fieldValues.AddNumber(v);
    }

    partial void OnBooleanParsed(bool v, bool? u, bool? f, string? c, uint? cp)
    {
        var fieldValues = GetFieldValues();
        fieldValues.AddBoolean(v);
    }

    partial void OnErrorParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b)
    {
        var fieldValues = GetFieldValues();
        if (!XLErrorParser.TryParseError(v, out var error))
            throw PartStructureException.InvalidAttributeFormat();

        fieldValues.AddError(error);
    }

    partial void OnStringParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b)
    {
        var fieldValues = GetFieldValues();
        fieldValues.AddString(v);
    }

    partial void OnDateTimeParsed(DateTime v, bool? u, bool? f, string? c, uint? cp)
    {
        var fieldValues = GetFieldValues();
        fieldValues.AddDateTime(v);
    }

    partial void OnIndexParsed(uint v)
    {
        var fieldValues = GetFieldValues();
        if (v >= fieldValues.SharedCount)
            throw PartStructureException.InvalidAttributeValue();

        fieldValues.AddIndex(v);
    }

    private XLPivotCacheValues GetFieldValues()
    {
        if (_fieldIdx >= _pivotCache.FieldCount)
            throw PartStructureException.IncorrectElementsCount();

        return _pivotCache.GetFieldValues(_fieldIdx++);
    }
}
