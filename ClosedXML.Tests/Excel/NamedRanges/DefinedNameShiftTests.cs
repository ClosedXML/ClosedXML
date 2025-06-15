using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.NamedRanges;

internal class DefinedNameShiftTests
{
    [TestCaseSource(nameof(GetShiftTestCases))]
    public void Defined_name_shifts_references_in_formula_on_structural_changes(Action<IXLWorksheet> shiftAction, string expectedNameFormula)
    {
        using var wb = new XLWorkbook();
        var shiftSheet = wb.AddWorksheet("Shift");
        var otherSheet = wb.AddWorksheet("Other");

        var workbookName = wb.DefinedNames.Add("WorkbookName", "Shift!$C$3:$D$5");
        var shiftName = shiftSheet.DefinedNames.Add("ShiftName", "Shift!$C$3:$D$5");
        var otherName = otherSheet.DefinedNames.Add("OtherName", "Other!$C$3:$D$5");

        shiftAction(shiftSheet);

        Assert.AreEqual(expectedNameFormula, workbookName.RefersTo);
        Assert.AreEqual(expectedNameFormula, shiftName.RefersTo);
        Assert.AreEqual("Other!$C$3:$D$5", otherName.RefersTo);
    }

    [Test]
    public void Shift_references_inside_formulas()
    {
        // When a shift happens, defined name will keep the original formula, references inside
        // the formula will be shifted though. Issue #2713.
        using var wb = new XLWorkbook();
        var lookupSheet = wb.AddWorksheet("Lookup");
        var name = wb.DefinedNames.Add("Name", "OFFSET(Lookup!$G$2,0,0,COUNTA(Lookup!$G:$G)-1,1)");

        lookupSheet.Row(1).Delete();

        Assert.AreEqual("OFFSET(Lookup!$G$1,0,0,COUNTA(Lookup!$G:$G)-1,1)", name.RefersTo);
    }

    public static IEnumerable<object[]> GetShiftTestCases
    {
        get
        {
            // Insert and shift reference down
            yield return Test(ws => ws.Range("C2:D2").InsertRowsBelow(4), "Shift!$C$7:$D$9");

            // Insert and shift reference right
            yield return Test(ws => ws.Range("B2:C7").InsertColumnsBefore(4), "Shift!$G$3:$H$5");

            // Delete area and shift cells up. Also partially shrinks reference to 1-row from 3-rows original.
            yield return Test(ws => ws.Range("C2:D4").Delete(XLShiftDeletedCells.ShiftCellsUp), "Shift!$C$2:$D$2");

            // Delete area and shift cells left. Also partially shrinks reference to 1-column from 2-column original.
            yield return Test(ws => ws.Range("B2:C7").Delete(XLShiftDeletedCells.ShiftCellsLeft), "Shift!$B$3:$B$5");

            // When the whole area is deleted, the defined name is a #REF! error.
            yield return Test(ws => ws.Range("A1:E8").Delete(XLShiftDeletedCells.ShiftCellsLeft), "#REF!");

            // When insert area doesn't push reference (in this case columns are different), no change will happen.
            yield return Test(ws => ws.Range("A1:B2").InsertRowsBelow(4), "Shift!$C$3:$D$5");

            // When insert area would cause a split, keep original reference.
            yield return Test(ws => ws.Range("C2").Delete(XLShiftDeletedCells.ShiftCellsUp), "Shift!$C$3:$D$5");

            yield break;

            static object[] Test(Action<IXLWorksheet> shift, string result)
            {
                return new object[] { shift, result };
            }
        }
    }
}
