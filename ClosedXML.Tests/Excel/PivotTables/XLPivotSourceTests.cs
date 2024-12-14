using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables;

/// <summary>
/// Tests for classes that implement <c>IXLPivotSource</c>.
/// </summary>
[TestFixture]
internal class XLPivotSourceTests
{
    [Test]
    public void Can_load_and_save_all_source_types()
    {
        // Teh test files contains all possible pivot cache sources. The output is mangled, but
        // Excel can open it and use refresh on each pivot table. External workbook is in the same
        // directory: PivotTable-AllSources-external-data.xlsx
        // TODO:Test file currently doesn't contain consolidate and scenario cache source. Will be added in subsequent PR
        TestHelper.LoadSaveAndCompare(
            @"Other\PivotTable\Sources\PivotTable-AllSources-input.xlsx",
            @"Other\PivotTable\Sources\PivotTable-AllSources-output.xlsx");
    }
}
