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
            ce.RegisterFunction("MINA", 1, int.MaxValue, MinA, FunctionFlags.Range, AllowRange.All);
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
            ce.RegisterFunction("STDEVA", 1, int.MaxValue, StDevA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEVP", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEVPA", 1, int.MaxValue, StDevPA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEV.S", 1, int.MaxValue, StDev, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("STDEV.P", 1, int.MaxValue, StDevP, FunctionFlags.Range, AllowRange.All);
            //STEYX	Returns the standard error of the predicted y-value for each x in the regression
            //TDIST	Returns the Student's t-distribution
            //TINV	Returns the inverse of the Student's t-distribution
            //TREND	Returns values along a linear trend
            //TRIMMEAN	Returns the mean of the interior of a data set
            //TTEST	Returns the probability associated with a Student's t-test
            ce.RegisterFunction("VAR", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VARA", 1, int.MaxValue, VarA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VARP", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VARPA", 1, int.MaxValue, VarPA, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VAR.S", 1, int.MaxValue, Var, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("VAR.P", 1, int.MaxValue, VarP, FunctionFlags.Range, AllowRange.All);
            //WEIBULL	Returns the Weibull distribution
            //ZTEST	Returns the one-tailed probability-value of a z-test
        }

        private static AnyValue Average(CalcContext ctx, Span<AnyValue> args)
        {
            return Average(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue Average(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            if (!tally.Tally(ctx, args, new SumState()).TryPickT0(out var state, out var error))
                return error;

            if (state.Count == 0)
                return XLError.DivisionByZero;

            return state.Sum / state.Count;
        }

        private static AnyValue AverageA(CalcContext ctx, Span<AnyValue> args)
        {
            return Average(ctx, args, TallyAll.WithArrayText);
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
            return Count(ctx, args, TallyNumbers.IgnoreErrors);
        }

        internal static AnyValue Count(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (args.Length < 1)
                return XLError.IncompatibleValue;

            var result = tally.Tally(ctx, args, new CountState(0));
            if (!result.TryPickT0(out var state, out var error))
                return error;

            return state.Count;
        }

        private static AnyValue CountA(CalcContext ctx, Span<AnyValue> args)
        {
            return Count(ctx, args, TallyAll.IncludeErrors);
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
                if (criteriaRanges.All(criteriaPair => CalcEngineHelpers.ValueSatisfiesCriteria(criteriaPair.Item2[i], criteriaPair.Item1, ce)))
                    count++;

                processedCount++;
            }

            // Add count of empty cells outside the used range if they match criteria
            if (criteriaRanges.All(criteriaPair => CalcEngineHelpers.ValueSatisfiesCriteria(string.Empty, criteriaPair.Item1, ce)))
            {
                count += (totalCount - processedCount);
            }

            // done
            return count;
        }

        private static AnyValue DevSq(CalcContext ctx, Span<AnyValue> args)
        {
            var result = GetSquareDiffSum(ctx, args, TallyNumbers.Default);
            if (!result.TryPickT0(out var squareDiff, out var error))
                return error;

            // An outlier, most others return #DIV/0! when they can't calculate mean.
            if (squareDiff.Count == 0)
                return XLError.NumberInvalid;

            return squareDiff.Sum;
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
            var tally = TallyNumbers.Default.Tally(ctx, args, new LogSumState(0.0, 0));
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
            return Max(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue Max(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            var result = tally.Tally(ctx, args, new MaxState());
            if (!result.TryPickT0(out var state, out var error))
                return error;

            if (!state.HasValues)
                return 0;

            return state.Max;
        }

        private static AnyValue MaxA(CalcContext ctx, Span<AnyValue> args)
        {
            return Max(ctx, args, TallyAll.Default);
        }

        private static AnyValue Median(CalcContext ctx, Span<AnyValue> args)
        {
            // There is a better median algorithm that uses two heaps, but NetFx
            // doesn't have heap structure.
            var result = TallyNumbers.Default.Tally(ctx, args, new ValuesState(new List<double>()));
            if (!result.TryPickT0(out var state, out var error))
                return error;

            var allNumbers = state.Values;
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
            return Min(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue Min(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            var result = tally.Tally(ctx, args, new MinState());

            if (!result.TryPickT0(out var state, out var error))
                return error;

            // Not even one non-ignored value found, return 0.
            if (!state.HasValues)
                return 0;

            return state.Min;
        }

        private static AnyValue MinA(CalcContext ctx, Span<AnyValue> args)
        {
            return Min(ctx, args, TallyAll.Default);
        }

        private static AnyValue StDev(CalcContext ctx, Span<AnyValue> args)
        {
            return StDev(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue StDev(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
                return error;

            if (squareDiff.Count <= 1)
                return XLError.DivisionByZero;

            return Math.Sqrt(squareDiff.Sum / (squareDiff.Count - 1));
        }

        private static AnyValue StDevA(CalcContext ctx, Span<AnyValue> args)
        {
            return StDev(ctx, args, TallyAll.Default);
        }

        private static AnyValue StDevP(CalcContext ctx, Span<AnyValue> args)
        {
            return StDevP(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue StDevP(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
                return error;

            if (squareDiff.Count < 1)
                return XLError.DivisionByZero;

            return Math.Sqrt(squareDiff.Sum / squareDiff.Count);
        }

        private static AnyValue StDevPA(CalcContext ctx, Span<AnyValue> args)
        {
            return StDevP(ctx, args, TallyAll.Default);
        }

        private static AnyValue Var(CalcContext ctx, Span<AnyValue> args)
        {
            return Var(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue Var(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
                return error;

            if (squareDiff.Count <= 1)
                return XLError.DivisionByZero;

            return squareDiff.Sum / (squareDiff.Count - 1);
        }

        private static AnyValue VarA(CalcContext ctx, Span<AnyValue> args)
        {
            return Var(ctx, args,TallyAll.Default);
        }

        private static AnyValue VarP(CalcContext ctx, Span<AnyValue> args)
        {
            return VarP(ctx, args, TallyNumbers.Default);
        }

        internal static AnyValue VarP(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (!GetSquareDiffSum(ctx, args, tally).TryPickT0(out var squareDiff, out var error))
                return error;

            if (squareDiff.Count < 1)
                return XLError.DivisionByZero;

            return squareDiff.Sum / squareDiff.Count;
        }

        private static AnyValue VarPA(CalcContext ctx, Span<AnyValue> args)
        {
            return VarP(ctx, args, TallyAll.Default);
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

        /// <summary>
        /// Calculate <c>SUM((x_i - mean_x)^2)</c> and number of samples. This method uses two-pass algorithm.
        /// There are several one-pass algorithms, but they are not numerically stable. In this case, accuracy
        /// takes precedence (plus VAR/STDEV are not a very frequently used function). Excel might have used
        /// those one-pass formulas in the past (see <em>Statistical flaws in Excel</em>), but doesn't seem to
        /// be using them anymore.
        /// </summary>
        private static OneOf<SquareDiff, XLError> GetSquareDiffSum(CalcContext ctx, Span<AnyValue> args, ITally tally)
        {
            if (!tally.Tally(ctx, args, new SumState(0.0, 0)).TryPickT0(out var sumState, out var sumError))
                return sumError;

            if (sumState.Count == 0)
                return new SquareDiff(0.0, 0, double.NaN);

            var sampleMean = sumState.Sum / sumState.Count;

            // Calculate sum of squares of deviations from sample mean
            var initialSquareDiffState = new SquareDiff(Sum: 0.0, Count: 0, SampleMean: sampleMean);
            var result = tally.Tally(ctx, args, initialSquareDiffState);

            if (!result.TryPickT0(out var squareDiff, out var error))
                return error;

            return squareDiff;
        }

        private readonly record struct SumState(double Sum, int Count) : ITallyState<SumState>
        {
            public SumState Tally(double number) => new(Sum + number, Count + 1);
        }

        private readonly record struct SquareDiff(double Sum, int Count, double SampleMean) : ITallyState<SquareDiff>
        {
            public SquareDiff Tally(double sampleValue)
            {
                var diff = sampleValue - SampleMean;
                var sum = Sum + diff * diff;
                return new SquareDiff(sum, Count + 1, SampleMean);
            }
        }

        private readonly record struct MinState(double Min, bool HasValues) : ITallyState<MinState>
        {
            public MinState() : this(double.MaxValue, false)
            {
            }

            public MinState Tally(double number) => new(Math.Min(Min, number), true);
        }

        private readonly record struct MaxState(double Max, bool HasValues) : ITallyState<MaxState>
        {
            public MaxState() : this(double.MinValue, false)
            {
            }

            public MaxState Tally(double number) => new(Math.Max(Max, number), true);
        }

        private readonly record struct LogSumState(double LogSum, int Count) : ITallyState<LogSumState>
        {
            public LogSumState Tally(double number)
            {
                var logSum = LogSum + Math.Log(number);
                return new(logSum, Count + 1);
            }
        }

        private readonly record struct ValuesState(List<double> Values) : ITallyState<ValuesState>
        {
            public ValuesState Tally(double number)
            {
                Values.Add(number);
                return new ValuesState(Values);
            }
        }

        private readonly record struct CountState(int Count) : ITallyState<CountState>
        {
            public CountState Tally(double number) => new(Count + 1);
        }
    }
}
