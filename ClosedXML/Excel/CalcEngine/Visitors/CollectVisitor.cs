using System;
using System.Collections.Generic;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine.Visitors;

internal abstract class CollectVisitor<TContext> : IAstFactory<object?, object?, TContext>
{
    public virtual object? LogicalValue(TContext context, SymbolRange range, bool value)
    {
        return default;
    }

    public virtual object? NumberValue(TContext context, SymbolRange range, double value)
    {
        return default;
    }

    public virtual object? TextValue(TContext context, SymbolRange range, string text)
    {
        return default;
    }

    public virtual object? ErrorValue(TContext context, SymbolRange range, ReadOnlySpan<char> error)
    {
        return default;
    }

    public virtual object? ArrayNode(TContext context, SymbolRange range, int rows, int columns, IReadOnlyList<object?> elements)
    {
        return default;
    }

    public virtual object? BlankNode(TContext context, SymbolRange range)
    {
        return default;
    }

    public virtual object? LogicalNode(TContext context, SymbolRange range, bool value)
    {
        return default;
    }

    public virtual object? ErrorNode(TContext context, SymbolRange range, ReadOnlySpan<char> error)
    {
        return default;
    }

    public virtual object? NumberNode(TContext context, SymbolRange range, double value)
    {
        return default;
    }

    public virtual object? TextNode(TContext context, SymbolRange range, string text)
    {
        return default;
    }

    public virtual object? Reference(TContext context, SymbolRange range, ReferenceArea reference)
    {
        return default;
    }

    public virtual object? SheetReference(TContext context, SymbolRange range, string sheet, ReferenceArea reference)
    {
        return default;
    }

    public virtual object? Reference3D(TContext context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        return default;
    }

    public virtual object? ExternalSheetReference(TContext context, SymbolRange range, int workbookIndex, string sheet,
        ReferenceArea reference)
    {
        return default;
    }

    public virtual object? ExternalReference3D(TContext context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet,
        ReferenceArea reference)
    {
        return default;
    }

    public virtual object? Function(TContext context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<object?> arguments)
    {
        return default;
    }

    public virtual object? Function(TContext context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<object?> args)
    {
        return default;
    }

    public virtual object? ExternalFunction(TContext context, SymbolRange range, int workbookIndex, string sheetName,
        ReadOnlySpan<char> functionName, IReadOnlyList<object?> arguments)
    {
        return default;
    }

    public virtual object? ExternalFunction(TContext context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName,
        IReadOnlyList<object?> arguments)
    {
        return default;
    }

    public virtual object? CellFunction(TContext context, SymbolRange range, RowCol cell, IReadOnlyList<object?> arguments)
    {
        return default;
    }

    public virtual object? StructureReference(TContext context, SymbolRange range, StructuredReferenceArea area, string? firstColumn,
        string? lastColumn)
    {
        return default;
    }

    public virtual object? StructureReference(TContext context, SymbolRange range, string table, StructuredReferenceArea area,
        string? firstColumn, string? lastColumn)
    {
        return default;
    }

    public virtual object? ExternalStructureReference(TContext context, SymbolRange range, int workbookIndex, string table,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return default;
    }

    public virtual object? Name(TContext context, SymbolRange range, string name)
    {
        return default;
    }

    public virtual object? SheetName(TContext context, SymbolRange range, string sheet, string name)
    {
        return default;
    }

    public virtual object? ExternalName(TContext context, SymbolRange range, int workbookIndex, string name)
    {
        return default;
    }

    public virtual object? ExternalSheetName(TContext context, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        return default;
    }

    public virtual object? BinaryNode(TContext context, SymbolRange range, BinaryOperation operation, object? leftNode, object? rightNode)
    {
        return default;
    }

    public virtual object? Unary(TContext context, SymbolRange range, UnaryOperation operation, object? node)
    {
        return default;
    }

    public virtual object? Nested(TContext context, SymbolRange range, object? node)
    {
        return default;
    }
}
