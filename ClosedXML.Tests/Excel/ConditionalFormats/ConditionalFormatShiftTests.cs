using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.ConditionalFormats;

[TestFixture]
public class ConditionalFormatShiftTests
{
    [Test]
    public void CFShiftedOnColumnInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).InsertColumnsAfter(2);
        var cf = ws.ConditionalFormats.ToArray();

        Assert.AreEqual(5, cf.Length);
        Assert.AreEqual("A1:A1", cf[0].Ranges.ToSpaceList());
        Assert.AreEqual("A2:D2", cf[1].Ranges.ToSpaceList());
        Assert.AreEqual("A3:E3", cf[2].Ranges.ToSpaceList());
        Assert.AreEqual("B4:D6", cf[3].Ranges.ToSpaceList());
        Assert.AreEqual("E7:F7", cf[4].Ranges.ToSpaceList());
    }

    [Test]
    public void CFShiftedOnRowInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).InsertRowsBelow(2);
        var cf = ws.ConditionalFormats.ToArray();

        Assert.AreEqual(5, cf.Length);
        Assert.AreEqual("A1:A1", cf[0].Ranges.ToSpaceList());
        Assert.AreEqual("B1:B4", cf[1].Ranges.ToSpaceList());
        Assert.AreEqual("C1:C5", cf[2].Ranges.ToSpaceList());
        Assert.AreEqual("D2:F4", cf[3].Ranges.ToSpaceList());
        Assert.AreEqual("G6:G7", cf[4].Ranges.ToSpaceList());
    }

    [Test]
    public void CFShiftedOnColumnDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).Delete();
        var cf = ws.ConditionalFormats.ToArray();

        Assert.AreEqual(4, cf.Length);
        Assert.AreEqual("A1:A1", cf[0].Ranges.ToSpaceList());
        Assert.AreEqual("A2:A2", cf[1].Ranges.ToSpaceList());
        Assert.AreEqual("A3:B3", cf[2].Ranges.ToSpaceList());
        Assert.AreEqual("B7:C7", cf[3].Ranges.ToSpaceList());
    }

    [Test]
    public void CFShiftedOnRowDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).Delete();
        var cf = ws.ConditionalFormats.ToArray();

        Assert.AreEqual(4, cf.Length);
        Assert.AreEqual("A1:A1", cf[0].Ranges.ToSpaceList());
        Assert.AreEqual("B1:B1", cf[1].Ranges.ToSpaceList());
        Assert.AreEqual("C1:C2", cf[2].Ranges.ToSpaceList());
        Assert.AreEqual("G3:G4", cf[3].Ranges.ToSpaceList());
    }

    [Test]
    public void CFShiftedTruncateRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.AsRange().AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Red);
        var cf = ws.ConditionalFormats.Single();

        ws.Row(2).InsertRowsAbove(1);
        Assert.IsTrue(cf.Range.RangeAddress.IsValid);
        Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", cf.Ranges.ToSpaceList());

        ws.Column(2).InsertColumnsAfter(1);
        Assert.IsTrue(cf.Range.RangeAddress.IsValid);
        Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", cf.Ranges.ToSpaceList());
    }

    [TestCaseSource(nameof(GetSplitTestCases))]
    public void Conditional_formats_can_split_when_inserted_or_deleted_doesnt_shift_whole_cf_area(string initialCfArea, Action<IXLWorksheet> shiftAction, string shiftedCfArea)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range(initialCfArea).AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Red);

        shiftAction(ws);

        Assert.AreEqual(shiftedCfArea, ws.ConditionalFormats.Single().Ranges.ToSpaceList());
    }

    [TestCaseSource(nameof(GetFormulaShiftTestCases))]
    public void Conditional_format_shifts_formulas(string cfArea, string formula, Action<IXLWorksheet> shiftAction, string shiftedFormula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Range(cfArea).AddConditionalFormat().WhenEquals(formula).Fill.SetBackgroundColor(XLColor.Red);

        shiftAction(ws);

        Assert.AreEqual(shiftedFormula, ws.ConditionalFormats.Single().Values.Single().Value.Value);
    }

    [Test]
    public void Conditional_format_shifts_formula_only_if_formula_cells_are_affected_by_shift()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").AddConditionalFormat().WhenEquals("=A1").Fill.SetBackgroundColor(XLColor.Red);

        // Shift only range with CF, not the referenced cells. Formula retains same A1 address
        ws.Range("B1").InsertRowsAbove(1);

        Assert.AreEqual("A1", ws.ConditionalFormats.Single().Values.Single().Value.Value);
        Assert.AreEqual("B2:B2", ws.ConditionalFormats.Single().Ranges.ToSpaceList());
    }

    [TestCase("A1")]
    [TestCase("Sheet!A1")]
    public void Conditional_format_does_not_shift_values_that_look_like_formulas(string formulaLookAlike)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("B1").AddConditionalFormat().WhenEquals(formulaLookAlike).Fill.SetBackgroundColor(XLColor.Red);

        ws.Row(1).InsertRowsAbove(1);

        Assert.AreEqual(formulaLookAlike, ws.ConditionalFormats.Single().Values.Single().Value.Value);
    }

    [Test]
    public void Conditional_format_formula_referencing_another_sheet_is_updated_when_another_sheet_is_modified()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet");
        var another = wb.AddWorksheet("Another");
        sheet.Cell("A1").AddConditionalFormat().WhenEquals("=Another!B2").Fill.SetBackgroundColor(XLColor.Red);

        another.Row(1).InsertRowsAbove(2);

        Assert.AreEqual("Another!B4", sheet.ConditionalFormats.Single().Values.Single().Value.Value);
    }

    private static IEnumerable<TestCaseData<string, Action<IXLWorksheet>, string>> GetSplitTestCases()
    {
        yield return new TestCaseData<string, Action<IXLWorksheet>, string>("A1:C3", ws => ws.Range("A1").InsertRowsAbove(2), "B1:C3 A3:A5");
        yield return new TestCaseData<string, Action<IXLWorksheet>, string>("A1:C3", ws => ws.Range("A2").InsertColumnsBefore(2), "A1:C1 C2:E2 A3:C3");
        yield return new TestCaseData<string, Action<IXLWorksheet>, string>("A10:C12", ws => ws.Range("B10").Delete(XLShiftDeletedCells.ShiftCellsUp), "A10:A12 B10:B11 C10:C12");
        yield return new TestCaseData<string, Action<IXLWorksheet>, string>("D1:G3", ws => ws.Range("A1:C2").Delete(XLShiftDeletedCells.ShiftCellsLeft), "A1:D2 D3:G3");
    }

    private static IEnumerable<TestCaseData<string, string, Action<IXLWorksheet>, string>> GetFormulaShiftTestCases()
    {
        // Smoke test
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("A1", "=B1", ws => ws.Row(1).InsertRowsAbove(1), "B2");

        // Insert or delete whole row or column
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Row(1).InsertRowsAbove(1), "IF(C4>A2,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Row(2).InsertRowsAbove(1), "IF(C4>A1,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Column(2).InsertColumnsBefore(2), "IF(E3>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Column(1).InsertColumnsBefore(1), "IF(D3>B1,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Row(2).Delete(), "IF(C2>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Row(1).Delete(), "IF(C2>#REF!,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Column(2).Delete(), "IF(B3>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Column(1).Delete(), "IF(B3>#REF!,TRUE,FALSE)");

        // Insert or delete area that shifts only portion of formula
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("B2:C2").InsertRowsAbove(2), "IF(C5>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("A1").InsertRowsAbove(2), "IF(C3>A3,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("B2:B3").InsertColumnsBefore(2), "IF(E3>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("A1").InsertColumnsBefore(2), "IF(C3>C1,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("B2:C2").Delete(XLShiftDeletedCells.ShiftCellsUp), "IF(C2>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("A1").Delete(XLShiftDeletedCells.ShiftCellsUp), "IF(C3>#REF!,TRUE,FALSE)");

        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("B1:B3").Delete(XLShiftDeletedCells.ShiftCellsLeft), "IF(B3>A1,TRUE,FALSE)");
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("C3:E5", "=IF(C3>A1,TRUE,FALSE)", ws => ws.Range("A1:A2").Delete(XLShiftDeletedCells.ShiftCellsLeft), "IF(C3>#REF!,TRUE,FALSE)");

        // Do not shift formulas form other sheets
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("A2", "=Other!A2", ws => ws.Row(1).InsertRowsAbove(10), "Other!A2");

        // Shift formulas with sheet name of the sheet with CF
        yield return new TestCaseData<string, string, Action<IXLWorksheet>, string>("A2", "=Sheet!B2", ws => ws.Row(1).InsertRowsAbove(1), "Sheet!B3");

    }
}
