using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine.Visitors;

internal static class FormulaTransformation
{
    private static readonly Lazy<PrefixTree> FutureFunctionSet = new(() => PrefixTree.Build(XLConstants.FutureFunctionMap.Value.Keys));

    private static readonly RenameFunctionsVisitor RemapFutureFunctions = new(XLConstants.FutureFunctionMap);

    /// <summary>
    /// Add necessary prefixes to a user-supplied future functions without a prefix (e.g.
    /// <c>acot(A5)/2</c> to <c>_xlfn.ACOT(A5)/2</c>).
    /// </summary>
    internal static string FixFutureFunctions(string formula, string sheetName, XLSheetPoint origin)
    {
        // A preliminary check that formula might contain future function. There are two reasons to do this first:
        // * Although parsing is relatively cheap, it's not free. Checking for string is far cheaper.
        // * Risk management, parser might fail for some formulas and limit fallout in such case.
        if (!MightContainFutureFunction(formula.AsSpan()))
            return formula;

        return FormulaConverter.ModifyA1(formula, sheetName, origin.Row, origin.Column, RemapFutureFunctions);
    }

    private static bool MightContainFutureFunction(ReadOnlySpan<char> formula)
    {
        for (var i = 0; i < formula.Length; ++i)
        {
            if (FutureFunctionSet.Value.IsPrefixOf(formula[i..]))
                return true;
        }

        return false;
    }

    /// <summary>
    /// All functions must have chars in the <c>.</c>-<c>_</c> range (trie range).
    /// </summary>
    private readonly record struct PrefixTree
    {
        private const char LowestChar = '.';
        private const char HighestChar = '_';

        /// <summary>
        /// Indicates the node represents a full prefix. Leaves are always ends and middle nodes
        /// sometimes (e.g. AB and ABC).
        /// </summary>
        private bool IsEnd { get; init; }

        /// <summary>
        /// Something transitions to this tree.
        /// </summary>
        [MemberNotNullWhen(false, nameof(Transitions))]
        private bool IsLeaf => Transitions is null;

        /// <summary>
        /// Index is a character minus <see cref="LowestChar"/>. The possible range of characters
        /// is from <see cref="LowestChar"/> to <see cref="HighestChar"/>.
        /// </summary>
        private PrefixTree[]? Transitions { get; init; }

        public static PrefixTree Build(IEnumerable<string> names)
        {
            var root = new PrefixTree { Transitions = new PrefixTree[HighestChar - LowestChar + 1] };
            foreach (var name in names)
                root.Insert(name.AsSpan());

            return root;
        }

        public bool IsPrefixOf(ReadOnlySpan<char> text)
        {
            var current = this;
            foreach (var c in text)
            {
                if (current.IsEnd)
                    return true;

                if (current.Transitions is null)
                    return false;

                var upperChar = char.ToUpperInvariant(c);
                if (upperChar is < LowestChar or > HighestChar)
                    return false;

                current = current.Transitions[upperChar - LowestChar];
            }

            return current.IsEnd;
        }

        private void Insert(ReadOnlySpan<char> functionName)
        {
            // Prev is necessary to update previous list due to immutability
            Debug.Assert(functionName.Length > 0);
            var prevTransitions = System.Array.Empty<PrefixTree>();
            var prevIndex = -1;
            var curNode = this;
            foreach (var c in functionName)
            {
                // All future function names are uppercase and in range, no need to transform.
                var transitionIndex = c - LowestChar;
                if (curNode.IsLeaf)
                {
                    // Current node is a leaf and thus has no transitions. Add them (kind of complicated thanks to readonly struct).
                    var currentTransitions = new PrefixTree[HighestChar - LowestChar + 1];
                    prevTransitions[prevIndex] = prevTransitions[prevIndex] with { Transitions = currentTransitions };
                    prevTransitions = currentTransitions;

                    // Move along the to a new node
                    curNode = currentTransitions[transitionIndex];
                }
                else
                {
                    prevTransitions = curNode.Transitions;
                    curNode = curNode.Transitions[transitionIndex];
                }

                prevIndex = transitionIndex;
            }

            prevTransitions[prevIndex] = prevTransitions[prevIndex] with { IsEnd = true };
        }
    }
}
