using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace ClosedXML.Excel.CalcEngine
{
    internal class Tally : IEnumerable<Object>
    {
        private readonly List<object> _list = new List<object>();
        private readonly bool NumbersOnly;

        private double[] _numericValues;


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

            NumbersOnly = numbersOnly;
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
                _numericValues = null;
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

            _numericValues = null;
        }

        public void AddValue(Object v)
        {
            _list.Add(v);
            _numericValues = null;
        }

        public double Count()
        {
            return Count(NumbersOnly);
        }

        public double Count(bool numbersOnly)
        {
            if (numbersOnly)
                return NumericValuesInternal().Length;
            else
                return _list.Count(o => !Statistical.IsBlank(o));
        }

        IEnumerable<double> NumericValuesEnumerable()
        {
            foreach (var value in _list)
            {
                double tmp;
                var vEnumerable = value as IEnumerable;
                if (vEnumerable == null)
                {
                    if (double.TryParse(value.ToString(), out tmp))
                        yield return tmp;
                }
                else
                {
                    foreach (var v in vEnumerable)
                    {
                        if (double.TryParse(v.ToString(), out tmp))
                            yield return tmp;
                        break;
                    }
                }
            }
        }

        double[] NumericValuesInternal()
            => LazyInitializer.EnsureInitialized(ref _numericValues, () => NumericValuesEnumerable().ToArray());

        public IEnumerable<double> NumericValues()
            => NumericValuesInternal().AsEnumerable();

        public double Product()
        {
            var nums = NumericValuesInternal();
            return nums.Length == 0
                ? 0
                : nums.Aggregate(1d, (a, b) => a * b);
        }

        public double Sum() => NumericValuesInternal().Sum();

        public double Average()
        {
            var nums = NumericValuesInternal();
            if (nums.Length == 0) throw new ApplicationException("No values");
            return nums.Average();
        }

        public double Min()
        {
            var nums = NumericValuesInternal();
            return nums.Length == 0 ? 0 : nums.Min();
        }

        public double Max()
        {
            var nums = NumericValuesInternal();
            return nums.Length == 0 ? 0 : nums.Max();
        }

        public double Range() => Max() - Min();

        static double Sum2(IEnumerable<double> nums)
        {
            return nums.Sum(d => d * d);
        }

        public double VarP()
        {
            var nums = NumericValuesInternal();
            var avg = nums.Average();
            var sum2 = Sum2(nums);
            var count = nums.Length;
            return count <= 1 ? 0 : sum2 / count - avg * avg;
        }

        public double StdP()
        {
            var nums = NumericValuesInternal();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            var count = nums.Length;
            return count <= 1 ? 0 : Math.Sqrt(sum2 / count - avg * avg);
        }

        public double Var()
        {
            var nums = NumericValuesInternal();
            var avg = nums.Average();
            var sum2 = Sum2(nums);
            var count = nums.Length;
            return count <= 1 ? 0 : (sum2 / count - avg * avg) * count / (count - 1);
        }

        public double Std()
        {
            var values = NumericValuesInternal();
            var count = values.Length;
            double ret = 0;
            if (count != 0)
            {
                //Compute the Average
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2
                double sum = values.Sum(d => Math.Pow(d - avg, 2));
                //Put it all together
                ret = Math.Sqrt((sum) / (count - 1));
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
