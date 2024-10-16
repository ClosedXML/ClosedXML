using System;

namespace ClosedXML.Excel.CalcEngine.Functions;

internal class TallyNumbers : ITally
{
    private readonly bool _ignoreScalarBlank;

    /// <summary>
    /// Tally numbers.
    /// </summary>
    internal static readonly TallyNumbers Default = new();

    /// <summary>
    /// Ignore blank from scalar values. Basically used for <c>PRODUCT</c> function, so it doesn't end up with 0.
    /// </summary>
    internal static readonly TallyNumbers WithoutScalarBlank = new(true);

    private TallyNumbers(bool ignoreScalarBlank = false)
    {
        _ignoreScalarBlank = ignoreScalarBlank;
    }

    /// <summary>
    /// The method tries to convert scalar arguments to numbers, but ignores non-numbers in
    /// reference/array. Any error found is propagated to the result.
    /// </summary>
    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        var tally = initialState;
        foreach (var arg in args)
        {
            if (arg.TryPickScalar(out var scalar, out var collection))
            {
                if (_ignoreScalarBlank && scalar.IsBlank)
                    continue;

                // Scalars are converted to number.
                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                    return error;

                tally = tally.Tally(number);
            }
            else
            {
                var valuesIterator = !collection.TryPickT0(out var array, out var reference)
                    ? ctx.GetNonBlankValues(reference)
                    : array;
                foreach (var value in valuesIterator)
                {
                    if (value.TryPickError(out var error))
                        return error;

                    // For arrays and references, only the number type is used. Other types are ignored.
                    if (value.TryPickNumber(out var number))
                        tally = tally.Tally(number);
                }
            }
        }

        return tally;
    }
}
