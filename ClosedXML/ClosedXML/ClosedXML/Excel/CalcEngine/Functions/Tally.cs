using System;
using System.Linq;
using System.Collections.Generic;
using System.Net;
using System.Collections;

namespace ClosedXML.Excel.CalcEngine
{
    internal class Tally
    {
        private readonly List<object> _list = new List<object>();

        public Tally(){}
        public Tally(IEnumerable<Expression> p)
        {
            foreach (var e in p)
            {
                Add(e);
            }
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
            _list.Add(e.Evaluate());
        }

        public void AddValue(Object v)
        {
            _list.Add(v);
        }

        public double Count() { return _list.Count; }
        public double CountA()
        {
            Double cntA = 0;
            foreach (var value in _list)
            {
                var strVal = value as String;
                if (value != null && (strVal == null || !XLHelper.IsNullOrWhiteSpace(strVal)))
                    cntA++;
            }
            return cntA;
        }
        public List<Double> Numerics()
        {
            List<Double> retVal = new List<double>();
            foreach (var value in _list)
            {
                Double tmp;
                if (Double.TryParse(value.ToString(), out tmp))
                {
                    retVal.Add(tmp);
                }
            }
            return retVal;
        }
        public double Product()
        {
            var nums = Numerics();
            if (nums.Count == 0) return 0;

            Double retVal = 1;
            nums.ForEach(n => retVal *= n);

            return retVal;
        }
        public double Sum() { return Numerics().Sum(); }
        public double Average()
        {
            return Numerics().Average();
        }
        public double Min() { return Numerics().Min(); }
        public double Max() { return Numerics().Max(); }
        public double Range()
        {
            var nums = Numerics();
            return nums.Max() - nums.Min();
        }

        private double Sum2(List<Double> nums)
        {
            return nums.Sum(d => d * d);
        }

        public double VarP()
        {
            var nums = Numerics();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count <= 1 ? 0 : sum2 / nums.Count - avg * avg;
        }
        public double StdP()
        {
            var nums = Numerics();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count <= 1 ? 0 : Math.Sqrt(sum2 / nums.Count - avg * avg);
        }
        public double Var()
        {
            var nums = Numerics();
            var avg = nums.Average();
            var sum2 = nums.Sum(d => d * d);
            return nums.Count <= 1 ? 0 : (sum2 / nums.Count - avg * avg) * nums.Count / (nums.Count - 1);
        }
        public double Std()
        {
            var values = Numerics();
            double ret = 0;
            if (values.Count > 0)
            {
                //Compute the Average      
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));
                //Put it all together      
                ret = Math.Sqrt((sum) / (values.Count() - 1));
            }
            return ret;

        }
    }
}
