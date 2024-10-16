using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine.Functions;

/// <summary>
/// A tally function for *A functions (e.g. AverageA, MinA, MaxA). The behavior is buggy in Excel,
/// because they doesn't count logical values in array, but do count them in reference ¯\_(ツ)_/¯.
///
/// <list type="bullet">
///   <item>Scalar values are converted to number, conversion might lead to errors.</item>
///   <item>Array values ignore logical and text (unless <c>ignoreArrayText</c> is <c>false</c>).</item>
///   <item>Reference values include logical, text is evaluated as zero.</item>
/// </list>
/// Any error is propagated.
/// </summary>
internal class TallyAll : ITally
{
    private readonly bool _ignoreArrayText;
    private readonly bool _includeErrors;

    /// <inheritdoc cref="TallyAll"/>
    /// <remarks>This tally ignores text in arrays.</remarks>
    internal static readonly ITally Default = new TallyAll(ignoreArrayText: true);

    /// <inheritdoc cref="TallyAll"/>
    /// <remarks>This tally counts text in arrays as <c>0</c>.</remarks>
    internal static readonly ITally WithArrayText = new TallyAll(ignoreArrayText: false);

    /// <summary>
    /// Include errors as number 0.
    /// </summary>
    internal static readonly ITally IncludeErrors = new TallyAll(includeErrors: true);

    private TallyAll(bool ignoreArrayText = true, bool includeErrors = false)
    {
        _ignoreArrayText = ignoreArrayText;
        _includeErrors = includeErrors;
    }

    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        var state = initialState;
        foreach (var arg in args)
        {
            if (arg.TryPickScalar(out var scalar, out var collection))
            {
                // Scalars are converted to number.
                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                {
                    if (!_includeErrors)
                        return error;

                    number = 0;
                }

                // All scalars are counted
                state = state.Tally(number);
            }
            else
            {
                bool isArray;
                IEnumerable<ScalarValue> valuesIterator;
                if (collection.TryPickT0(out var array, out var reference))
                {
                    valuesIterator = array;
                    isArray = true;
                }
                else
                {
                    valuesIterator = ctx.GetNonBlankValues(reference);
                    isArray = false;
                }
                foreach (var value in valuesIterator)
                {
                    // Blank lines are ignored. Logical are counted in reference, but not in array.
                    if (!isArray && value.TryPickLogical(out var logical))
                    {
                        state = state.Tally(logical ? 1 : 0);
                    }
                    else if (value.TryPickNumber(out var number))
                    {
                        state = state.Tally(number);
                    }
                    else if (value.IsText && (!isArray || !_ignoreArrayText))
                    {
                        // Some *A functions consider text in an array (e.g. {"3", "Hello"}) as a zero and others don't.
                        // The text values from cells behave differently. Unlike array, the *A functions consider cell text as 0.
                        state = state.Tally(0);
                    }
                    else if (value.TryPickError(out var error))
                    {
                        if (!_includeErrors)
                            return error;

                        state = state.Tally(0);
                    }
                }
            }
        }

        return state;
    }
}
