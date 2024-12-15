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
        // The test files contains all possible pivot cache sources. The output is mangled, but
        // Excel can open it and use refresh on each pivot table. External workbook is in the same
        // directory: PivotTable-AllSources-external-data.xlsx.
        // The pivot table that uses connection has a connection to the external workbook
        // PivotTable-AllSources-external-data.xlsx. The connection uses an absolute path, so it
        // needs to be updated according to real directory. Doesn't affect CI, because connection
        // is not actually used to get data.
        // Scenario doesn't throw on refresh, but it incomplete. The cache source is correct though.
        //
        // Open the workbook and click Pivot Table Analyze - Refresh - Refresh All. It shouldn't
        // report an error.
        TestHelper.LoadSaveAndCompare(
            @"Other\PivotTable\Sources\PivotTable-AllSources-input.xlsx",
            @"Other\PivotTable\Sources\PivotTable-AllSources-output.xlsx");
    }
}
