using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Functions;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Statistical
    {
        public static void Register(FunctionRegistry ce)
        {
            //ce.RegisterFunction("AVEDEV", AveDev, 1, int.MaxValue);
            ce.RegisterFunction("AVERAGE", 1, int.MaxValue, Average, FunctionFlags.Range, AllowRange.All); // Returns the average (arithmetic mean) of the arguments
            ce.RegisterFunction("AVERAGEA", 1, int.MaxValue, AverageA, FunctionFlags.Range, AllowRange.All);
            //BETADIST	Returns the beta cumulative distribution function
            //BETAINV   Returns the inverse of the cumulative distribution function for a specified beta distribution
            ce.RegisterFunction("BINOMDIST", 4, 4, Adapt(BinomDist), FunctionFlags.Scalar); //BINOMDIST	Returns the individual term binomial distribution probability
            ce.RegisterFunction("BINOM.DIST", 4, 4, Adapt(BinomDist), FunctionFlags.Scalar); // In theory more precise BINOMDIST.
            //CHIDIST	Returns the one-tailed probability of the chi-squared distribution
            //CHIINV	Returns the inverse of the one-tailed probability of the chi-squared distribution
            //CHITEST	Returns the test for independence
            //CONFIDENCE	Returns the confidence interval for a population mean
            //CORREL	Returns the correlation coefficient between two data sets
            ce.RegisterFunction("COUNT", 1, int.MaxValue, Count, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("COUNTA", 1, 255, CountA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("COUNTBLANK", 1, CountBlank, AllowRange.All);
            ce.RegisterFunction("COUNTIF", 2, CountIf, AllowRange.Only, 0);
            ce.RegisterFunction("COUNTIFS", 2, 255, CountIfs, AllowRange.Only, Enumerable.Range(0, 128).Select(x => x * 2).ToArray());
            //COVAR	Returns covariance, the average of the products of paired deviations
            //CRITBINOM	Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
            ce.RegisterFunction("DEVSQ", 1, 255, DevSq, FunctionFlags.Range, AllowRange.All); // Returns the sum of squares of deviations
            //EXPONDIST	Returns the exponential distribution
            //FDIST	Returns the F probability distribution
            //FINV	Returns the inverse of the F probability distribution
            ce.RegisterFunction("FISHER", 1, Fisher); // Returns the Fisher transformation
            //FISHERINV	Returns the inverse of the Fisher transformation
            //FORECAST	Returns a value along a linear trend
            //FREQUENCY	Returns a frequency distribution as a vertical array
            //FTEST	Returns the result of an F-test
            //GAMMADIST	Returns the gamma distribution
            //GAMMAINV	Returns the inverse of the gamma cumulative distribution
            //GAMMALN	Returns the natural logarithm of the gamma function, Γ(x)
            ce.RegisterFunction("GEOMEAN", 1, 255, GeoMean, FunctionFlags.Range, AllowRange.All); // Returns the geometric mean
            //GROWTH	Returns values along an exponential trend
            //HARMEAN	Returns the harmonic mean
            //HYPGEOMDIST	Returns the hypergeometric distribution
            //INTERCEPT	Returns the intercept of the linear regression line
            //KURT	Returns the kurtosis of a data set
            //LARGE	Returns the k-th largest value in a data set
            ce.RegisterFunction("LARGE", 2, 2, Adapt(Large), FunctionFlags.Range, AllowRange.Only, 0);
            //LINEST	Returns the parameters of a linear trend
            //LOGEST	Returns the parameters of an exponential trend
            //LOGINV	Returns the inverse of the lognormal distribution
            //LOGNORMDIST	Returns the cumulative lognormal distribution
            ce.RegisterFunction("MAX", 1, 255, Max, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MAXA", 1, int.MaxValue, MaxA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MEDIAN", 1, int.MaxValue, Median, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MIN", 1, int.MaxValue, Min, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MINA", 1, int.MaxValue, MinA, AllowRange.All);
            //MODE	Returns the most common value in a data set
            //NEGBINOMDIST	Returns the negative binomial distribution
            //NORMDIST	Returns the normal cumulative distribution
            //NORMINV	Returns the inverse of the normal cumulative distribution
            //NORMSDIST	Returns the standard normal cumulative distribution
            //NORMSINV	Returns the inverse of the standard normal cumulative distribution
            //PEARSON	Returns the Pearson product moment correlation coefficient
            //PERCENTILE	Returns the k-th percentile of values in a range
            //PERCENTRANK	Returns the percentage rank of a value in a data set
            //PERMUT	Returns the number of permutations for a given number of objects
            //POISSON	Returns the Poisson distribution
            //PROB	Returns the probability that values in a range are between two limits
            //QUARTILE	Returns the quartile of a data set
            //RANK	Returns the rank of a number in a list of numbers
            //RSQ	Returns the square of the Pearson product moment correlation coefficient
            //SKEW	Returns the skewness of a distribution
            //SLOPE	Returns the slope of the linear regression line
            //SMALL	Returns the k-th smallest value in a data set
            //STANDARDIZE	Returns a normalized value
            ce.RegisterFunction("STDEV", 1, int.MaxValue, StDev, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEVA", 1, int.MaxValue, StDevA, AllowRange.All);
            ce.RegisterFunction("STDEVP", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEVPA", 1, int.MaxValue, StDevPA, AllowRange.All);
            ce.RegisterFunction("STDEV.S", 1, int.MaxValue, StDev, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEV.P", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
            //STEYX	Returns the standard error of the predicted y-value for each x in the regression
            //TDIST	Returns the Student's t-distribution
            //TINV	Returns the inverse of the Student's t-distribution
            //TREND	Returns values along a linear trend
            //TRIMMEAN	Returns the mean of the interior of a data set
            //TTEST	Returns the probability associated with a Student's t-test
            ce.RegisterFunction("VAR", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VARA", 1, int.MaxValue, VarA, AllowRange.All);
            ce.RegisterFunction("VARP", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VARPA", 1, int.MaxValue, VarPA, AllowRange.All);
            ce.RegisterFunction("VAR.S", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VAR.P", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
            //WEIBULL	Returns the Weibull distribution
            //ZTEST	Returns the one-tailed probability-value of a z-test
        }

        private static AnyValue Average(CalcContext ctx, Span<AnyValue> args)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            var result = TallyNumbers(ctx, args, (Sum: 0.0, Count: 0), static (state, number) => (state.Sum + number, state.Count + 1));
            if (!result.TryPickT0(out var state, out var error))
                return error;

            if (state.Count == 0)
                return XLError.DivisionByZero;

            return state.Sum / state.Count;
        }

        private static AnyValue AverageA(CalcContext ctx, Span<AnyValue> args)
        {
            var state = (Sum: 0.0, Count: 0);
            var result = TallyAll(ctx, args, state, static (state, value) => (state.Sum + value, state.Count + 1));

            if (!result.TryPickT0(out var tally, out var error))
                return error;

            if (tally.Count == 0)
                return XLError.DivisionByZero;

            return tally.Sum / tally.Count;
        }

        private static AnyValue BinomDist(double numberSuccesses, double numberTrials, double successProbability, bool cumulativeFlag)
        {
            if (successProbability is < 0 or > 1)
                return XLError.NumberInvalid;

            if (cumulativeFlag)
            {
                var cdf = 0d;
                for (var y = 0; y <= numberSuccesses; ++y)
                {
                    var result = BinomDist(y, numberTrials, successProbability);
                    if (!result.TryPickT0(out var pf, out var error))
                        return error;

                    cdf += pf;
                }

                if (double.IsNaN(cdf) || double.IsInfinity(cdf))
                    return XLError.NumberInvalid;

                return cdf;
            }
            else
            {
                var result = BinomDist(numberSuccesses, numberTrials, successProbability);
                if (!result.TryPickT0(out var binomDist, out var error))
                    return error;

                return binomDist;
            }
        }

        private static OneOf<double, XLError> BinomDist(double x, double n, double p)
        {
            if (!XLMath.CombinChecked(n, x).TryPickT0(out var combinations, out var error))
                return error;

            x = Math.Floor(x);
            n = Math.Floor(n);
            var binomDist = combinations * Math.Pow(p, x) * Math.Pow(1 - p, n - x);
            if (double.IsNaN(binomDist) || double.IsInfinity(binomDist))
                return XLError.NumberInvalid;

            return binomDist;
        }

        private static AnyValue Count(CalcContext ctx, Span<AnyValue> args)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            var count = 0;
            foreach (var arg in args)
            {
                if (arg.TryPickScalar(out var scalar, out var collection))
                {
                    // Scalars are converted to number.
                    if (scalar.ToNumber(ctx.Culture).TryPickT0(out _, out _))
                        count++;
                }
                else
                {
                    var valuesIterator = collection.TryPickT0(out var array, out var reference)
                        ? array
                        : ctx.GetNonBlankValues(reference);
                    foreach (var value in valuesIterator)
                    {
                        // For arrays and references, only the number type is used. Other types are ignored.
                        if (value.TryPickNumber(out var number))
                            count++;
                    }
                }
            }

            return count;
        }

        private static AnyValue CountA(CalcContext ctx, Span<AnyValue> values)
        {
            var result = values.Aggregate(
                ctx,
                initialValue: 0,
                noElementsResult: 0,
                collectionFilter: value =>
                {
                    // Blanks in collections (i.e. references, because arrays can't contain blanks)
                    // are not counted and thus are filtered out.
                    if (value.IsBlank)
                        return false;

                    // Everything else is counted, including errors.
                    return true;
                },
                // Any scalar value (including errors, including blank, if is passed directly as
                // an argument) is counted as one non-empty element.
                convert: (_, _) => 1,
                aggregate: static (acc, cur) => acc + cur);

            if (!result.TryPickT0(out var nonEmptyCount, out var error))
                return error;

            return nonEmptyCount;
        }

        private static object CountBlank(List<Expression> p)
        {
            if ((p[0] as XObjectExpression)?.Value as CellRangeReference == null)
                return XLError.NoValueAvailable;

            var e = (XObjectExpression)p[0];
            long totalCount = CalcEngineHelpers.GetTotalCellsCount(e);
            long nonBlankCount = 0;
            foreach (var value in e)
            {
                if (!CalcEngineHelpers.ValueIsBlank(value))
                    nonBlankCount++;
            }

            return 0d + totalCount - nonBlankCount;
        }

        private static object CountIf(List<Expression> p)
        {
            XLCalcEngine ce = new XLCalcEngine(CultureInfo.CurrentCulture);
            var cnt = 0.0;
            long processedCount = 0;
            if (p[0] is XObjectExpression ienum)
            {
                long totalCount = CalcEngineHelpers.GetTotalCellsCount(ienum);
                var criteria = p[1].Evaluate();
                foreach (var value in ienum)
                {
                    if (CalcEngineHelpers.ValueSatisfiesCriteria(value, criteria, ce))
                        cnt++;
                    processedCount++;
                }

                // Add count of empty cells outside the used range if they match criteria
                if (CalcEngineHelpers.ValueSatisfiesCriteria(string.Empty, criteria, ce))
                    cnt += (totalCount - processedCount);
            }

            return cnt;
        }

        private static object CountIfs(List<Expression> p)
        {
            // get parameters
            var ce = new XLCalcEngine(CultureInfo.CurrentCulture);
            long count = 0;

            int numberOfCriteria = p.Count / 2;

            long totalCount = 0;
            // prepare criteria-parameters:
            var criteriaRanges = new Tuple<object, List<object>>[numberOfCriteria];
            for (int criteriaPair = 0; criteriaPair < numberOfCriteria; criteriaPair++)
            {
                var criteriaRange = (XObjectExpression)p[criteriaPair * 2];
                var criterion = p[(criteriaPair * 2) + 1].Evaluate();
                var criteriaRangeValues = new List<object>();
                foreach (var value in criteriaRange)
                {
                    criteriaRangeValues.Add(value);
                }

                criteriaRanges[criteriaPair] = new Tuple<object, List<object>>(
                    criterion,
                    criteriaRangeValues);

                if (totalCount == 0)
                    totalCount = CalcEngineHelpers.GetTotalCellsCount(criteriaRange);
            }

            long processedCount = 0;
            for (var i = 0; i < criteriaRanges[0].Item2.Count; i++)
            {
                if (criteriaRanges.All(criteriaPair => CalcEngineHelpers.ValueSatisfiesCriteria(
                                                       criteriaPair.Item2[i], criteriaPair.Item1, ce)))
                    count++;

                processedCount++;
            }

            // Add count of empty cells outside the used range if they match criteria
            if (criteriaRanges.All(criteriaPair => CalcEngineHelpers.ValueSatisfiesCriteria(
                                                   string.Empty, criteriaPair.Item1, ce)))
            {
                count += (totalCount - processedCount);
            }

            // done
            return count;
        }

        private static AnyValue DevSq(CalcContext ctx, Span<AnyValue> args)
        {
            var result = GetSquareDiffSum(ctx, args);
            if (!result.TryPickT0(out var state, out var error))
                return error;

            // An outlier, most others return #DIV/0! when they can't calculate mean.
            if (state.Count == 0)
                return XLError.NumberInvalid;

            return state.SquareDiffSum;
        }

        private static object Fisher(List<Expression> p)
        {
            var x = (double)p[0];
            if (x <= -1 || x >= 1) return XLError.NumberInvalid;

            return 0.5 * Math.Log((1 + x) / (1 - x));
        }

        private static AnyValue GeoMean(CalcContext ctx, Span<AnyValue> args)
        {
            // Rather than interrupting a cycle early, just add it all
            // go through all values anyway. I don't want to code same
            // loop 1000 times and non-positive numbers will be rare.
            var state = (LogSum: 0.0, Count: 0);
            var tally = TallyNumbers(ctx, args, state, static (state, itemValue) =>
            {
                var logSum = state.LogSum + Math.Log(itemValue);
                return (logSum, state.Count + 1);
            });

            if (!tally.TryPickT0(out var geoMean, out var error))
                return error;

            if (geoMean.Count == 0)
                return XLError.NumberInvalid;

            // Some value was negative or zero. NaN plus whatever is NaN, infinity
            // plus whatever is also infinity.
            if (double.IsInfinity(geoMean.LogSum) || double.IsNaN(geoMean.LogSum))
                return XLError.NumberInvalid;

            return Math.Exp(geoMean.LogSum / geoMean.Count);
        }

        private static AnyValue Max(CalcContext ctx, Span<AnyValue> args)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            var state = (Max: double.MinValue, HasValues: false);
            var result = TallyNumbers(ctx, args, state, static (state, value) => (Math.Max(state.Max, value), true));

            if (!result.TryPickT0(out var tally, out var error))
                return error;

            if (!tally.HasValues)
                return 0;

            return tally.Max;
        }

        private static AnyValue MaxA(CalcContext ctx, Span<AnyValue> args)
        {
            // Text in array is skipped, in reference counted as 0
            var state = (Max: double.MinValue, HasValues: false);
            var result = TallyAll(ctx, args, state, static (state, value) => (Math.Max(state.Max, value), true), countArrayTextAsZero: false);

            if (!result.TryPickT0(out var tally, out var error))
                return error;

            // Not even one non-ignored value found, return 0.
            if (!tally.HasValues)
                return 0;

            return tally.Max;
        }

        private static AnyValue Median(CalcContext ctx, Span<AnyValue> args)
        {
            // There is a better median algorithm that uses two heaps, but NetFx
            // doesn't have heap structure.
            var list = new List<double>();
            var result = TallyNumbers(ctx, args, list, static (list, number) =>
            {
                list.Add(number);
                return list;
            });

            if (!result.TryPickT0(out var allNumbers, out var error))
                return error;

            if (allNumbers.Count == 0)
                return XLError.NumberInvalid;

            allNumbers.Sort();

            var halfIndex = allNumbers.Count / 2;
            var hasEvenCount = allNumbers.Count % 2 == 0;
            if (hasEvenCount)
                return (allNumbers[halfIndex - 1] + allNumbers[halfIndex]) / 2;

            return allNumbers[halfIndex];
        }

        private static AnyValue Min(CalcContext ctx, Span<AnyValue> args)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            double? min = null;
            var result = TallyNumbers(ctx, args, min, static (min, itemValue) => min.HasValue ? Math.Min(min.Value, itemValue) : itemValue);

            return result.Match<AnyValue>(m => m ?? 0, e => e);
        }

        private static object MinA(List<Expression> p)
        {
            return GetTally(p).Min();
        }

        private static AnyValue StDev(CalcContext ctx, Span<AnyValue> args)
        {
            if (!GetSquareDiffSum(ctx, args).TryPickT0(out var tallied, out var error))
                return error;

            if (tallied.Count <= 1)
                return XLError.DivisionByZero;

            return Math.Sqrt(tallied.SquareDiffSum / (tallied.Count - 1));
        }

        private static object StDevA(List<Expression> p)
        {
            return GetTally(p).Std();
        }

        private static AnyValue StDevP(CalcContext ctx, Span<AnyValue> args)
        {
            if (!GetSquareDiffSum(ctx, args).TryPickT0(out var tallied, out var error))
                return error;

            if (tallied.Count < 1)
                return XLError.DivisionByZero;

            return Math.Sqrt(tallied.SquareDiffSum / tallied.Count);
        }

        private static object StDevPA(List<Expression> p)
        {
            return GetTally(p).StdP();
        }

        private static AnyValue Var(CalcContext ctx, Span<AnyValue> args)
        {
            if (!GetSquareDiffSum(ctx, args).TryPickT0(out var tally, out var error))
                return error;

            if (tally.Count <= 1)
                return XLError.DivisionByZero;

            return tally.SquareDiffSum / (tally.Count - 1);
        }

        private static object VarA(List<Expression> p)
        {
            return GetTally(p).Var();
        }

        private static AnyValue VarP(CalcContext ctx, Span<AnyValue> args)
        {
            if (!GetSquareDiffSum(ctx, args).TryPickT0(out var tally, out var error))
                return error;

            if (tally.Count < 1)
                return XLError.DivisionByZero;

            return tally.SquareDiffSum / tally.Count;
        }

        private static object VarPA(List<Expression> p)
        {
            return GetTally(p).VarP();
        }

        private static AnyValue Large(CalcContext ctx, AnyValue arrayParam, double kParam)
        {
            if (kParam < 1)
                return XLError.NumberInvalid;

            var k = (int)Math.Ceiling(kParam);

            IEnumerable<ScalarValue> values;
            int size;
            if (arrayParam.TryPickScalar(out var scalar, out var collection))
            {
                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                    return error;

                values = new ScalarValue[] { number };
                size = 1;
            }
            else if (collection.TryPickT0(out var array, out var reference))
            {
                values = array;
                size = array.Width * array.Height;
            }
            else
            {
                values = reference.GetCellsValues(ctx);
                size = reference.NumberOfCells;
            }

            // Pre-allocate array to reduce allocations during doubling of buffer.
            var total = new List<double>(size);
            foreach (var value in values)
            {
                if (value.IsError)
                    return value.GetError();

                if (value.IsNumber)
                    total.Add(value.GetNumber());
            }

            if (k > total.Count)
                return XLError.NumberInvalid;

            total.Sort();

            return total[^k];
        }

        // utility for tallying statistics
        private static Tally GetTally(List<Expression> p)
        {
            return new Tally(p);
        }

        /// <summary>
        /// Calculate <c>SUM((x_i - mean_x)^2)</c> and number of samples. This method uses two-pass algorithm.
        /// There are several one-pass algorithms, but they are not numerically stable. In this case, accuracy
        /// takes precedence (plus VAR/STDEV are not a very frequently used function). Excel might have used
        /// those one-pass formulas in the past (see <em>Statistical flaws in Excel</em>), but doesn't seem to
        /// be using them anymore.
        /// </summary>
        private static OneOf<(double SquareDiffSum, int Count), XLError> GetSquareDiffSum(CalcContext ctx, Span<AnyValue> args)
        {
            // Calculate mean
            var sumResult = TallyNumbers(ctx, args, (Sum: 0.0, Count: 0), static (state, number) => (state.Sum + number, state.Count + 1));
            if (!sumResult.TryPickT0(out var averageState, out var sumError))
                return sumError;

            if (averageState.Count == 0)
                return (0.0, 0);

            var average = averageState.Sum / averageState.Count;

            // Calculate sum of squares of deviations from sample mean
            var components = (SquareDiffSum: 0.0, Count: 0, SampleMean: average);
            var result = TallyNumbers(ctx, args, components, static (stdDev, sampleValue) =>
            {
                var diff = sampleValue - stdDev.SampleMean;
                var sum = stdDev.SquareDiffSum + diff * diff;
                return (sum, stdDev.Count + 1, stdDev.SampleMean);
            });

            if (!result.TryPickT0(out var tallied, out var error))
                return error;

            return (tallied.SquareDiffSum, tallied.Count);
        }

        /// <summary>
        /// The method tries to convert scalar arguments to numbers, but ignores non-numbers in
        /// reference/array. Any error found is propagated to the result.
        /// </summary>
        internal static OneOf<T, XLError> TallyNumbers<T>(CalcContext ctx, Span<AnyValue> args, T initValue, Func<T, double, T> tallyFunc)
        {
            var tally = initValue;
            foreach (var arg in args)
            {
                if (arg.TryPickScalar(out var scalar, out var collection))
                {
                    // Scalars are converted to number.
                    if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                        return error;

                    tally = tallyFunc(tally, number);
                }
                else
                {
                    var valuesIterator = collection.TryPickT0(out var array, out var reference)
                        ? array
                        : ctx.GetNonBlankValues(reference);
                    foreach (var value in valuesIterator)
                    {
                        if (value.TryPickError(out var error))
                            return error;

                        // For arrays and references, only the number type is used. Other types are ignored.
                        if (value.TryPickNumber(out var number))
                            tally = tallyFunc(tally, number);
                    }
                }
            }

            return tally;
        }

        /// <summary>
        /// A tally function for *A functions (e.g. AverageA, MinA, MaxA). The behavior is buggy in Excel,
        /// because they doesn't count logical values in array, but do count them in reference ¯\_(ツ)_/¯
        /// </summary>
        /// <paramref name="countArrayTextAsZero">When there is a text value in an array, should it be tallied
        ///   as <c>0</c> (<c>true</c>), or should it be ignored (<c>false</c>).</paramref>
        private static OneOf<T, XLError> TallyAll<T>(CalcContext ctx, Span<AnyValue> args, T initialState, Func<T, double, T> tallyFunc, bool countArrayTextAsZero = true)
        {
            var state = initialState;
            foreach (var arg in args)
            {
                if (arg.TryPickScalar(out var scalar, out var collection))
                {
                    // Scalars are converted to number.
                    if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                        return error;

                    // All scalars are counted
                    state = tallyFunc(state, number);
                }
                else
                {
                    var isArray = collection.TryPickT0(out var array, out var reference);
                    var valuesIterator = isArray ? array! : ctx.GetNonBlankValues(reference!);
                    foreach (var value in valuesIterator)
                    {
                        // Blank lines are ignored. Logical are counted in reference, but not in array.
                        if (!isArray && value.TryPickLogical(out var logical))
                        {
                            state = tallyFunc(state, logical ? 1 : 0);
                        }
                        else if (value.TryPickNumber(out var number))
                        {
                            state = tallyFunc(state, number);
                        }
                        else if (value.IsText && (!isArray || countArrayTextAsZero))
                        {
                            // Some *A functions consider text in an array (e.g. {"3", "Hello"}) as a zero and others don't.
                            // The text values from cells behave differently. Unlike array, the *A functions consider cell text as 0.
                            state = tallyFunc(state, 0);
                        }
                        else if (value.TryPickError(out var error))
                        {
                            return error;
                        }
                    }
                }
            }

            return state;
        }
    }
}
