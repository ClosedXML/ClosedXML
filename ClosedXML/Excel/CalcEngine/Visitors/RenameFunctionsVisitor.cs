using System;
using System.Collections.Generic;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A visitor for <see cref="FormulaConverter"/> that maps one name of a function to another.
/// </summary>
internal class RenameFunctionsVisitor : RefModVisitor
{
    /// <summary>
    /// Case insensitive dictionary of function names.
    /// </summary>
    private readonly Lazy<IReadOnlyDictionary<string, string>> _functionMap;

    internal RenameFunctionsVisitor(Lazy<IReadOnlyDictionary<string, string>> functionMap)
    {
        _functionMap = functionMap;
    }

    protected override ReadOnlySpan<char> ModifyFunction(ModContext ctx, ReadOnlySpan<char> functionName)
    {
        if (_functionMap.Value.TryGetValue(functionName.ToString(), out var mapped))
        {
            return mapped.AsSpan();
        }

        return functionName;
    }
}
