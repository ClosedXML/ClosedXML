using ClosedXML.Parser;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A collection of all references in the book (not others) found in a formula.
/// Created by <see cref="CollectRefsFactory"/>.
/// </summary>
internal class FormulaReferences
{
    private readonly string _formula;

    private FormulaReferences(string formula)
    {
        _formula = formula;
    }

    /// <summary>
    /// Is there a <c>#REF!</c> anywhere in the formula?
    /// </summary>
    internal bool ContainsRefError { get; set; }

    /// <summary>
    /// Areas without a sheet found in the formula.
    /// </summary>
    internal HashSet<XLReference> References { get; } = new();

    /// <summary>
    /// Areas with a sheet found in the formula.
    /// </summary>
    internal HashSet<XLSheetReference> SheetReferences { get; } = new();

    internal HashSet<(string Table, string Column, string Symbol)> StructuredReferences { get; } = new();

    internal static FormulaReferences ForFormula(string formula)
    {
        var references = new FormulaReferences(formula);
        FormulaParser<object?, object?, FormulaReferences>.CellFormulaA1(formula, references, CollectRefsFactory.Instance);
        return references;
    }

    internal bool ContainsSheet(string worksheetName)
    {
        return SheetReferences.Any(x => XLHelper.SheetComparer.Equals(x.Sheet, worksheetName));
    }

    internal XLRanges GetExternalRanges(XLWorkbook workbook, XLSheetPoint anchor)
    {
        var list = new XLRanges();
        foreach (var reference in SheetReferences)
        {
            if (workbook.TryGetWorksheet(reference.Sheet, out XLWorksheet sheet))
            {
                var rangeAddress = reference.Reference.ToRangeAddress(sheet, anchor);
                list.Add(sheet.Range(rangeAddress));
            }
        }

        foreach (var (tableName, column, _) in StructuredReferences)
        {
            if (workbook.TryGetTable(tableName, out var table))
                list.Add(table.DataRange.Column(column));
        }

        return list;
    }

    /// <summary>
    /// Factory to get all references (cells, tables, names) in local workbook.
    /// </summary>
    private class CollectRefsFactory : CollectVisitor<FormulaReferences>
    {
        public static readonly CollectRefsFactory Instance = new();

        public override object? ErrorNode(FormulaReferences context, SymbolRange range, ReadOnlySpan<char> error)
        {
            context.ContainsRefError = true;
            return base.ErrorNode(context, range, error);
        }

        public override object? Reference(FormulaReferences context, SymbolRange range, ReferenceArea reference)
        {
            context.References.Add(new XLReference(reference));
            return base.Reference(context, range, reference);
        }

        public override object? SheetReference(FormulaReferences context, SymbolRange range, string sheet, ReferenceArea reference)
        {
            context.SheetReferences.Add(new XLSheetReference(sheet, new XLReference(reference)));
            return base.SheetReference(context, range, sheet, reference);
        }

        public override object? StructureReference(FormulaReferences context, SymbolRange range, string table, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            // TODO: Temporary placeholder, extract range detection from CalculationVisitor
            if (firstColumn is not null)
                context.StructuredReferences.Add((table, firstColumn, context._formula.Substring(range.Start, range.End - range.Start)));

            return base.StructureReference(context, range, table, area, firstColumn, lastColumn);
        }
    }
}
