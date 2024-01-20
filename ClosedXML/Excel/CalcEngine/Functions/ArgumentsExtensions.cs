using System;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    /// <summary>
    /// An extension methods
    /// </summary>
    internal static class ArgumentsExtensions
    {
        /// <summary>
        /// Aggregate all values in the arguments of a function into a single value. If any value is error, return the error.
        /// </summary>
        /// <remarks>
        /// A lot of functions take all argument values and aggregate the values to a different value.
        /// These aggregation functions apply aggregation on each argument and if the argument is
        /// a collection (array/reference), the aggregation function is also applied to each element of
        /// the array/reference (e.g. <c>SUM({1, 2}, 3)</c> applies sum on each element of an array
        /// <c>{1,2}</c> and thus result is <c>1+2+3</c>).
        /// </remarks>
        /// <typeparam name="TValue">Type of the value that is being aggregated.</typeparam>
        /// <param name="args">Arguments of a function. Method goes over all elements of the arguments.</param>
        /// <param name="ctx">Calculation context.</param>>
        /// <param name="initialValue">
        /// Initial value of the accumulator. It is used as an input into the first call of <paramref name="aggregate"/>.
        /// </param>
        /// <param name="noElementsResult">
        /// What should be the result of aggregation, if there are no elements. Common choices are
        /// <see cref="XLError.IncompatibleValue"/> or the <paramref name="initialValue"/>.
        /// </param>
        /// <param name="aggregate">
        /// The aggregation function. First parameter is the accumulator, second parameter is the value of
        /// current element taken from <paramref name="convert"/>. Make sure the method is static lambda to
        /// avoid useless allocations.
        /// </param>
        /// <param name="convert">
        /// A function that converts a scalar value of an element into the <typeparamref name="TValue"/> or
        /// an error if it can't be converted. Make sure the method is static lambda to avoid useless allocations.
        /// </param>
        /// <param name="collectionFilter">
        /// Some functions skip elements in a array/reference that would be accepted as an argument,
        /// e.g. <c>SUM("1", {2,"4"})</c> is <c>3</c> - it converts string <c>"3"</c> to a number <c>3</c>
        /// in for root arguments, but omits element <c>"4"</c> in the array. This is a function that
        /// determines which elements to include and which to skip. If null, all elements of array are included and
        /// all values are treated same. Make sure the method is static lambda to avoid useless allocations.
        /// </param>
        public static OneOf<TValue, XLError> Aggregate<TValue>(
             this Span<AnyValue> args,
             CalcContext ctx,
             TValue initialValue,
             OneOf<TValue, XLError> noElementsResult,
             Func<TValue, TValue, TValue> aggregate,
             Func<ScalarValue, CalcContext, OneOf<TValue, XLError>> convert,
             Func<ScalarValue, bool>? collectionFilter = null)
        {
            var result = initialValue;
            var hasElement = false;
            foreach (var arg in args)
            {
                if (arg.TryPickScalar(out var scalar, out var collection))
                {
                    var conversionResult = convert(scalar, ctx);
                    if (!conversionResult.TryPickT0(out var elementValue, out var elementError))
                        return elementError;

                    hasElement = true;
                    result = aggregate(result, elementValue!);
                }
                else
                {
                    var valuesIterator = collection.TryPickT0(out var array, out var reference)
                        ? array!
                        : reference!.GetCellsValues(ctx);
                    foreach (var value in valuesIterator)
                    {
                        if (collectionFilter is not null && !collectionFilter(value))
                            continue;

                        var conversionResult = convert(value, ctx);
                        if (!conversionResult.TryPickT0(out var elementValue, out var elementError))
                            return elementError;

                        hasElement = true;
                        result = aggregate(result, elementValue!);
                    }
                }
            }

            return hasElement ? result : noElementsResult;
        }
    }
}
