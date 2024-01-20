using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Functions;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Statistical
    {
        public static void Register(FunctionRegistry ce)
        {
            //ce.RegisterFunction("AVEDEV", AveDev, 1, int.MaxValue);
            ce.RegisterFunction("AVERAGE", 1, int.MaxValue, Average, AllowRange.All); // Returns the average (arithmetic mean) of the arguments
            ce.RegisterFunction("AVERAGEA", 1, int.MaxValue, AverageA, AllowRange.All);
            //BETADIST	Returns the beta cumulative distribution function
            //BETAINV	Returns the inverse of the cumulative distribution function for a specified beta distribution
            //BINOMDIST	Returns the individual term binomial distribution probability
            //CHIDIST	Returns the one-tailed probability of the chi-squared distribution
            //CHIINV	Returns the inverse of the one-tailed probability of the chi-squared distribution
            //CHITEST	Returns the test for independence
            //CONFIDENCE	Returns the confidence interval for a population mean
            //CORREL	Returns the correlation coefficient between two data sets
            ce.RegisterFunction("COUNT", 1, int.MaxValue, Count, AllowRange.All);
            ce.RegisterFunction("COUNTA", 1, int.MaxValue, CountA, AllowRange.All);
            ce.RegisterFunction("COUNTBLANK", 1, CountBlank, AllowRange.All);
            ce.RegisterFunction("COUNTIF", 2, CountIf, AllowRange.Only, 0);
            ce.RegisterFunction("COUNTIFS", 2, 255, CountIfs, AllowRange.Only, Enumerable.Range(0, 128).Select(x => x * 2).ToArray());
            //COVAR	Returns covariance, the average of the products of paired deviations
            //CRITBINOM	Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
            ce.RegisterFunction("DEVSQ", 1, 255, DevSq, AllowRange.All); // Returns the sum of squares of deviations
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
            ce.RegisterFunction("GEOMEAN", 1, 255, Geomean, AllowRange.All); // Returns the geometric mean
            //GROWTH	Returns values along an exponential trend
            //HARMEAN	Returns the harmonic mean
            //HYPGEOMDIST	Returns the hypergeometric distribution
            //INTERCEPT	Returns the intercept of the linear regression line
            //KURT	Returns the kurtosis of a data set
            //LARGE	Returns the k-th largest value in a data set
            //LINEST	Returns the parameters of a linear trend
            //LOGEST	Returns the parameters of an exponential trend
            //LOGINV	Returns the inverse of the lognormal distribution
            //LOGNORMDIST	Returns the cumulative lognormal distribution
            ce.RegisterFunction("MAX", 1, 255, Max, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("MAXA", 1, int.MaxValue, MaxA, AllowRange.All);
            ce.RegisterFunction("MEDIAN", 1, int.MaxValue, Median, AllowRange.All);
            ce.RegisterFunction("MIN", 1, int.MaxValue, Min, AllowRange.All);
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
            ce.RegisterFunction("STDEV", 1, int.MaxValue, StDev, AllowRange.All);
            ce.RegisterFunction("STDEVA", 1, int.MaxValue, StDevA, AllowRange.All);
            ce.RegisterFunction("STDEVP", 1, int.MaxValue, StDevP, AllowRange.All);
            ce.RegisterFunction("STDEVPA", 1, int.MaxValue, StDevPA, AllowRange.All);
            ce.RegisterFunction("STDEV.S", 1, int.MaxValue, StDev);
            ce.RegisterFunction("STDEV.P", 1, int.MaxValue, StDevP);
            //STEYX	Returns the standard error of the predicted y-value for each x in the regression
            //TDIST	Returns the Student's t-distribution
            //TINV	Returns the inverse of the Student's t-distribution
            //TREND	Returns values along a linear trend
            //TRIMMEAN	Returns the mean of the interior of a data set
            //TTEST	Returns the probability associated with a Student's t-test
            ce.RegisterFunction("VAR", 1, int.MaxValue, Var, AllowRange.All);
            ce.RegisterFunction("VARA", 1, int.MaxValue, VarA, AllowRange.All);
            ce.RegisterFunction("VARP", 1, int.MaxValue, VarP, AllowRange.All);
            ce.RegisterFunction("VARPA", 1, int.MaxValue, VarPA, AllowRange.All);
            ce.RegisterFunction("VAR.S", 1, int.MaxValue, Var);
            ce.RegisterFunction("VAR.P", 1, int.MaxValue, VarP);
            //WEIBULL	Returns the Weibull distribution
            //ZTEST	Returns the one-tailed probability-value of a z-test
        }

        private static object Average(List<Expression> p)
        {
            return GetTally(p, true).Average();
        }

        private static object AverageA(List<Expression> p)
        {
            return GetTally(p, false).Average();
        }

        private static object Count(List<Expression> p)
        {
            return GetTally(p, true).Count();
        }

        private static object CountA(List<Expression> p)
        {
            return GetTally(p, false).Count();
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

        private static object DevSq(List<Expression> p)
        {
            return GetTally(p, true).DevSq();
        }

        private static object Fisher(List<Expression> p)
        {
            var x = (double)p[0];
            if (x <= -1 || x >= 1) return XLError.NumberInvalid;

            return 0.5 * Math.Log((1 + x) / (1 - x));
        }

        private static object Geomean(List<Expression> p)
        {
            return GetTally(p, true).GeoMean();
        }

        private static object Max(List<Expression> p)
        {
            return GetTally(p, true).Max();
        }

        private static AnyValue Max(CalcContext ctx, Span<AnyValue> args)
        {
            var result = args.Aggregate(
                ctx,
                initialValue: double.NegativeInfinity,
                noElementsResult: 0,
                aggregate: static (acc, current) => Math.Max(acc, current),
                convert: static (value, ctx) => value.ToNumber(ctx.Culture),
                collectionFilter: v => v.IsNumber || v.IsError); // Ignore blanks in references

            if (!result.TryPickT0(out var maximum, out var error))
                return error;

            if (double.IsInfinity(maximum))
                return 0;

            return maximum;
        }

        private static object MaxA(List<Expression> p)
        {
            return GetTally(p, false).Max();
        }

        private static object Median(List<Expression> p)
        {
            return GetTally(p, false).Median();
        }

        private static object Min(List<Expression> p)
        {
            return GetTally(p, true).Min();
        }

        private static object MinA(List<Expression> p)
        {
            return GetTally(p, false).Min();
        }

        private static object StDev(List<Expression> p)
        {
            return GetTally(p, true).Std();
        }

        private static object StDevA(List<Expression> p)
        {
            return GetTally(p, false).Std();
        }

        private static object StDevP(List<Expression> p)
        {
            return GetTally(p, true).StdP();
        }

        private static object StDevPA(List<Expression> p)
        {
            return GetTally(p, false).StdP();
        }

        private static object Var(List<Expression> p)
        {
            return GetTally(p, true).Var();
        }

        private static object VarA(List<Expression> p)
        {
            return GetTally(p, false).Var();
        }

        private static object VarP(List<Expression> p)
        {
            return GetTally(p, true).VarP();
        }

        private static object VarPA(List<Expression> p)
        {
            return GetTally(p, false).VarP();
        }

        // utility for tallying statistics
        private static Tally GetTally(List<Expression> p, bool numbersOnly)
        {
            return new Tally(p, numbersOnly);
        }
    }
}
