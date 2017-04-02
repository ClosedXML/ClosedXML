using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class Tally : IEnumerable<Object>
    {
        private readonly List<object> _list = new List<object>();
        private readonly bool NumbersOnly;

        public Tally()
            : this(false)
        { }

        public Tally(bool numbersOnly)
            : this(null, numbersOnly)
        { }

        public Tally(IEnumerable<Expression> p)
            : this(p, false)
        { }

        public Tally(IEnumerable<Expression> p, bool numbersOnly)
        {
            if (p != null)
            {
                foreach (var e in p)
                {
                    Add(e);
                }
            }

            this.NumbersOnly = numbersOnly;
        }

        public void Add(Expression e)
        {
            // handle enumerables
            var ienum = e as IEnumerable;
            if (ienum != null)
            {
                foreach (var value in ienum)
                {
                    _list.Add(value);
                }
                return;
            }

            // handle expressions
            var val = e.Evaluate();
            var valEnumerable = val as IEnumerable;
            if (valEnumerable == null || val is string)
                _list.Add(val);
            else
                foreach (var v in valEnumerable)
                    _list.Add(v);
        }

        public void AddValue(Object v)
        {
            _list.Add(v);
        }

        public double Count()
        {
            return this.Count(this.NumbersOnly);
        }

        public double Count(bool numbersOnly)
        {
            if (numbersOnly)
                return NumericValues().Count();
            else
                return _list.Where(o => !Statistical.IsBlank(o)).Count();
        }

        public IEnumerable<Double> NumericValues()
        {
            var retVal = new List<double>();
            foreach (var value in _list)
            {
                Double tmp;
                var vEnumerable = value as IEnumerable;
                if (vEnumerable == null && Double.TryParse(value.ToString(), out tmp))
                    yield return tmp;
                else
                {
                    foreach (var v in vEnumerable)
                    {
                        if (Double.TryParse(v.ToString(), out tmp))
                            yield return tmp;
                        break;
                    }
                }
            }
        }

        public double Product()
        {
            var nums = NumericValues();
            if (!nums.Any()) return 0;

            Double retVal = 1;
            nums.ForEach(n => retVal *= n);

            return retVal;
        }

        public double Sum() { return NumericValues().Sum(); }

        public double Average()
        {
            if (NumericValues().Any())
                return NumericValues().Average();
            else
                throw new ApplicationException("No values");
        }

        public double Min()
        {
            return NumericValues().Any() ? NumericValues().Min() : 0;
        }

        public double Max()
        {
            return NumericValues().Any() ? NumericValues().Max() : 0;
        }

        public double Range()
        {
            var nums = NumericValues();
            return nums.Max() - nums.Min();
        }

        private double Sum2(List<Double> nums)
        {
            return nums.Sum(d => d * d);
        }

        public double VarP()
        {
            var nums = NumericValues();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count() <= 1 ? 0 : sum2 / nums.Count() - avg * avg;
        }

        public double StdP()
        {
            var nums = NumericValues();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count() <= 1 ? 0 : Math.Sqrt(sum2 / nums.Count() - avg * avg);
        }

        public double Var()
        {
            var nums = NumericValues();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count() <= 1 ? 0 : (sum2 / nums.Count() - avg * avg) * nums.Count() / (nums.Count() - 1);
        }

        public double Std()
        {
            var values = NumericValues();
            double ret = 0;
            if (values.Any())
            {
                //Compute the Average
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2
                double sum = values.Sum(d => Math.Pow(d - avg, 2));
                //Put it all together
                ret = Math.Sqrt((sum) / (values.Count() - 1));
            }
            else
            {
                throw new ApplicationException("No values");
            }
            return ret;
        }

        public IEnumerator<object> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
