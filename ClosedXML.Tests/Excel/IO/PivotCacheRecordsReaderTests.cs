using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class PivotCacheRecordsReaderTests
{
    [Test]
    public void Can_read_all_record_item_types()
    {
        var sharedItems = new XLPivotCacheSharedItems();
        sharedItems.Add("First shared item");
        sharedItems.Add("Second shared item");

        ReadRecords(
            new[] { "Field 1" },
            $"""
             <pivotCacheRecords xmlns="{OpenXmlConst.Main2006SsNs}">
               <r>
                 <m/>
               </r>
               <r>
                 <n v="5.5"/>
               </r>
               <r>
                 <b v="true"/>
               </r>
               <r>
                 <e v="#NUM!"/>
               </r>
               <r>
                 <s v="Text"/>
               </r>
               <r>
                 <d v="2020-10-05"/>
               </r>
               <r>
                 <x v="1"/>
               </r>
             </pivotCacheRecords>
             """,
            (cache, reader) =>
            {
                reader.ReadRecordsToCache();

                var values = cache.GetFieldValues(0);
                Assert.That(values.GetCellValues(), Is.EquivalentTo(new XLCellValue[]
                {
                    Blank.Value,
                    5.5,
                    true,
                    XLError.NumberInvalid,
                    "Text",
                    new DateTime(2020, 10, 5),
                    "Second shared item"
                }));
            }, sharedItems);
    }

    [TestCase("<m/>")]
    [TestCase("<m/><m/><m/>")]
    public void All_records_must_have_same_number_of_items_as_there_is_cache_fields(string recordItems)
    {
        ReadRecords(
            new[] { "Field 1", "Field 2" },
            $"""
             <pivotCacheRecords xmlns="{OpenXmlConst.Main2006SsNs}">
               <r>{recordItems}</r>
             </pivotCacheRecords>
             """,
            (_, reader) =>
            {
                Assert.That(reader.ReadRecordsToCache, Throws
                    .Exception.TypeOf<PartStructureException>().And
                    .Message.StartsWith(PartStructureException.IncorrectElementsCount().Message));
            });
    }

    private static void ReadRecords(IReadOnlyList<string> fieldNames, string recordsXml, Action<XLPivotCache, PivotCacheRecordsReader> assert, XLPivotCacheSharedItems sharedItems = null)
    {
        using var wb = new XLWorkbook();
        var cache = wb.PivotCachesInternal.Add(new XLPivotSourceConnection(0));
        sharedItems ??= new XLPivotCacheSharedItems();
        foreach (var fieldName in fieldNames)
        {
            cache.AddCachedField(fieldName, new XLPivotCacheValues(sharedItems, new XLPivotCacheValuesStats()));
        }

        using var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(recordsXml));
        using var xmlTreeReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, false);
        var reader = new PivotCacheRecordsReader(xmlTreeReader, cache);
        assert(cache, reader);
    }
}
