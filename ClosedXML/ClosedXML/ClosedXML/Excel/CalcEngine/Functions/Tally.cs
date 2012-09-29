using System;
using System.Net;
using System.Collections;

namespace ClosedXML.Excel.CalcEngine
{
    internal class Tally
    {
        double _sum, _sum2, _cnt, _min, _max;
        bool _numbersOnly;

        public Tally(bool numbersOnly)
        {
            _numbersOnly = numbersOnly;
        }
        public Tally()
        {
        }

        public void Add(Expression e)
        {
            // handle enumerables
            var ienum = e as IEnumerable;
            if (ienum != null)
            {
                foreach (var value in ienum)
                {
                    AddValue(value);
                }
                return;
            }

            // handle expressions
            AddValue(e.Evaluate());
        }
        public void AddValue(object value)
        {
            // conversions
            if (!_numbersOnly)
            {
                // arguments that contain text evaluate as 0 (zero). 
                // empty text ("") evaluates as 0 (zero).
                if (value == null || value is string)
                {
                    value = 0;
                }
                // arguments that contain TRUE evaluate as 1; 
                // arguments that contain FALSE evaluate as 0 (zero).
                if (value is bool)
                {
                    value = (bool)value ? 1 : 0;
                }
            }

            // convert all numeric values to doubles
            if (value != null)
            {
                var typeCode = Type.GetTypeCode(value.GetType());
                if (typeCode >= TypeCode.Char && typeCode <= TypeCode.Decimal)
                {
                    value = Convert.ChangeType(value, typeof(double), System.Globalization.CultureInfo.CurrentCulture);
                }
            }

            // tally
            if (value is double)
            {
                var dbl = (double)value;
                _sum += dbl;
                _sum2 += dbl * dbl;
                _cnt++;
                if (_cnt == 1 || dbl < _min)
                {
                    _min = dbl;
                }
                if (_cnt == 1 || dbl > _max)
                {
                    _max = dbl;
                }
            }
        }
        public double Count() { return _cnt; }
        public double Sum() { return _sum; }
        public double Average() { return _sum / _cnt; }
        public double Min() { return _min; }
        public double Max() { return _max; }
        public double Range() { return _max - _min; }
        public double VarP()
        {
            var avg = Average();
            return _cnt <= 1 ? 0 : _sum2 / _cnt - avg * avg;
        }
        public double StdP()
        {
            var avg = Average();
            return _cnt <= 1 ? 0 : Math.Sqrt(_sum2 / _cnt - avg * avg);
        }
        public double Var()
        {
            var avg = Average();
            return _cnt <= 1 ? 0 : (_sum2 / _cnt - avg * avg) * _cnt / (_cnt - 1);
        }
        public double Std()
        {
            var avg = Average();
            return _cnt <= 1 ? 0 : Math.Sqrt((_sum2 / _cnt - avg * avg) * _cnt / (_cnt - 1));
        }
    }
}
