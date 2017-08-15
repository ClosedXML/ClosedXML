using System;
using System.Net;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Statistical
    {
        public static void Register(CalcEngine ce)
        {
            //ce.RegisterFunction("AVEDEV", AveDev, 1, int.MaxValue);
            ce.RegisterFunction("AVERAGE", 1, int.MaxValue, Average); // Returns the average (arithmetic mean) of the arguments
            ce.RegisterFunction("AVERAGEA", 1, int.MaxValue, AverageA);
            //BETADIST	Returns the beta cumulative distribution function
            //BETAINV	Returns the inverse of the cumulative distribution function for a specified beta distribution
            //BINOMDIST	Returns the individual term binomial distribution probability
            //CHIDIST	Returns the one-tailed probability of the chi-squared distribution
            //CHIINV	Returns the inverse of the one-tailed probability of the chi-squared distribution
            //CHITEST	Returns the test for independence
            //CONFIDENCE	Returns the confidence interval for a population mean
            //CORREL	Returns the correlation coefficient between two data sets
            ce.RegisterFunction("COUNT", 1, int.MaxValue, Count);
            ce.RegisterFunction("COUNTA", 1, int.MaxValue, CountA);
            ce.RegisterFunction("COUNTBLANK", 1, CountBlank);
            ce.RegisterFunction("COUNTIF", 2, CountIf);
            //COVAR	Returns covariance, the average of the products of paired deviations
            //CRITBINOM	Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
            //DEVSQ	Returns the sum of squares of deviations
            //EXPONDIST	Returns the exponential distribution
            //FDIST	Returns the F probability distribution
            //FINV	Returns the inverse of the F probability distribution
            //FISHER	Returns the Fisher transformation
            //FISHERINV	Returns the inverse of the Fisher transformation
            //FORECAST	Returns a value along a linear trend
            //FREQUENCY	Returns a frequency distribution as a vertical array
            //FTEST	Returns the result of an F-test
            //GAMMADIST	Returns the gamma distribution
            //GAMMAINV	Returns the inverse of the gamma cumulative distribution
            //GAMMALN	Returns the natural logarithm of the gamma function, Γ(x)
            //GEOMEAN	Returns the geometric mean
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
            ce.RegisterFunction("MAX", 1, int.MaxValue, Max);
            ce.RegisterFunction("MAXA", 1, int.MaxValue, MaxA);
            //MEDIAN	Returns the median of the given numbers
            ce.RegisterFunction("MIN", 1, int.MaxValue, Min);
            ce.RegisterFunction("MINA", 1, int.MaxValue, MinA);
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
            ce.RegisterFunction("STDEV", 1, int.MaxValue, StDev);
            ce.RegisterFunction("STDEVA", 1, int.MaxValue, StDevA);
            ce.RegisterFunction("STDEVP", 1, int.MaxValue, StDevP);
            ce.RegisterFunction("STDEVPA", 1, int.MaxValue, StDevPA);
            ce.RegisterFunction("STDEV.S", 1, int.MaxValue, StDev);
            ce.RegisterFunction("STDEV.P", 1, int.MaxValue, StDevP);
            //STEYX	Returns the standard error of the predicted y-value for each x in the regression
            //TDIST	Returns the Student's t-distribution
            //TINV	Returns the inverse of the Student's t-distribution
            //TREND	Returns values along a linear trend
            //TRIMMEAN	Returns the mean of the interior of a data set
            //TTEST	Returns the probability associated with a Student's t-test
            ce.RegisterFunction("VAR", 1, int.MaxValue, Var);
            ce.RegisterFunction("VARA", 1, int.MaxValue, VarA);
            ce.RegisterFunction("VARP", 1, int.MaxValue, VarP);
            ce.RegisterFunction("VARPA", 1, int.MaxValue, VarPA);
            ce.RegisterFunction("VAR.S", 1, int.MaxValue, Var);
            ce.RegisterFunction("VAR.P", 1, int.MaxValue, VarP);
            //WEIBULL	Returns the Weibull distribution
            //ZTEST	Returns the one-tailed probability-value of a z-test
        }



        static object Average(List<Expression> p)
        {
            return GetTally(p, true).Average();
        }
        static object AverageA(List<Expression> p)
        {
            return GetTally(p, false).Average();
        }
        static object Count(List<Expression> p)
        {
            return GetTally(p, true).Count();
        }
        static object CountA(List<Expression> p)
        {
            return GetTally(p, false).Count();
        }
        static object CountBlank(List<Expression> p)
        {
            var cnt = 0.0;
            foreach (Expression e in p)
            {
                var ienum = e as IEnumerable;
                if (ienum != null)
                {
                    foreach (var value in ienum)
                    {
                        if (IsBlank(value))
                            cnt++;
                    }
                }
                else
                {
                    if (IsBlank(e.Evaluate()))
                        cnt++;
                }
            }
            return cnt;
        }
        internal static bool IsBlank(object value)
        {
            return
                value == null ||
                value is string && ((string)value).Length == 0;
        }
        static object CountIf(List<Expression> p)
        {
            CalcEngine ce = new CalcEngine();
            var cnt = 0.0;
            var ienum = p[0] as IEnumerable;
            if (ienum != null)
            {
                var criteria = (string)p[1].Evaluate();
                foreach (var value in ienum)
                {
                    if (!IsBlank(value))
                    {
                        if (CalcEngineHelpers.ValueSatisfiesCriteria(value, criteria, ce))
                            cnt++;
                    }
                }
            }
            return cnt;
        }
        static object Max(List<Expression> p)
        {
            return GetTally(p, true).Max();
        }
        static object MaxA(List<Expression> p)
        {
            return GetTally(p, false).Max();
        }
        static object Min(List<Expression> p)
        {
            return GetTally(p, true).Min();
        }
        static object MinA(List<Expression> p)
        {
            return GetTally(p, false).Min();
        }
        static object StDev(List<Expression> p)
        {
            return GetTally(p, true).Std();
        }
        static object StDevA(List<Expression> p)
        {
            return GetTally(p, false).Std();
        }
        static object StDevP(List<Expression> p)
        {
            return GetTally(p, true).StdP();
        }
        static object StDevPA(List<Expression> p)
        {
            return GetTally(p, false).StdP();
        }
        static object Var(List<Expression> p)
        {
            return GetTally(p, true).Var();
        }
        static object VarA(List<Expression> p)
        {
            return GetTally(p, false).Var();
        }
        static object VarP(List<Expression> p)
        {
            return GetTally(p, true).VarP();
        }
        static object VarPA(List<Expression> p)
        {
            return GetTally(p, false).VarP();
        }

        // utility for tallying statistics
        static Tally GetTally(List<Expression> p, bool numbersOnly)
        {
            return new Tally(p, numbersOnly);
        }
    }
}
