using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.Exceptions;
using NUnit.Framework;
using System.Runtime.CompilerServices;

namespace ClosedXML.Tests.Excel.CalcEngine;

[TestFixture]
public class CalculationMemoryLimitTests
{
    [Test]
    public void All_arrays_in_formula_share_limit()
    {
        const int cells = 4;
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = cells * Unsafe.SizeOf<ScalarValue>() });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;

        ws.Cell("C1").SetFormulaA1($"A1:A{cells}+1");
        Assert.AreEqual(11.0, ws.Cell("C1").Value);

        ws.Cell("C2").SetFormulaA1($"(A1:A{cells}+1)+(B1:B{cells}+1)");
        Assert.Throws<LimitExceededException>(() => _ = ws.Cell("C2").Value);
    }

    [Test]
    public void Reference_to_cells_doesnt_not_instantiate_array_and_is_not_counted()
    {
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = 1 });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 4;
        ws.Cell("B1").SetFormulaA1("SUM(A1:A2)");

        Assert.AreEqual(7.0, ws.Cell("B1").Value);
    }

    [Test]
    public void Results_of_operations_are_counted()
    {
        // The multiplication operation will create an intermediate (and final in this case) result
        // that is an array.
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = 1 });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 4;
        ws.Cell("B1").SetFormulaA1("SUM(A1:A2 * 2)");

        Assert.Throws<LimitExceededException>(() => _ = ws.Cell("B1").Value);
    }

    [Test]
    public void Limit_checks_maximal_usage_at_any_time_of_calculation()
    {
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = 2 * Unsafe.SizeOf<ScalarValue>() });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 4;
        ws.Cell("B1").SetFormulaA1("SUM(A1:A2 * 2)");
        ws.Cell("B2").SetFormulaA1("SUM(A1:A2 * 2)");

        // The access to .Value causes recalculation of workbook and all cells are dirty.
        // But cells are calculated sequentially, so the memory buffer for each individual
        // cell calculation with its own formula doesn't exceed the limit.
        Assert.AreEqual(14.0, ws.Cell("B2").Value);
    }

    [Test]
    public void Expose_only_generic_closedxml_exception()
    {
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = 1 });
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").SetFormulaA1("A1:A2+1");
        Assert.That(() => _ = ws.Cell("B1").Value, Throws.InstanceOf<ClosedXMLException>());
        Assert.IsFalse(typeof(LimitExceededException).IsPublic);
    }

    [Test]
    public void Zero_means_no_limit()
    {
        using var wb = new XLWorkbook(new LoadOptions { MaxCalculationMemoryBytes = 0 });
        var ws = wb.AddWorksheet();
        ws.Range("A1:B100").Value = 1;
        ws.Cell("C1").SetFormulaA1("SUM(A1:B100+1)");
        Assert.AreEqual(400.0, ws.Cell("C1").Value);
    }

    [Test]
    public void Static_evaluate_methods_have_no_limit()
    {
        const string formula = "SUM({1,2}*{3;4})";
        const double expectedResult = 21.0;
        Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExprCurrent(formula));
    }
}
