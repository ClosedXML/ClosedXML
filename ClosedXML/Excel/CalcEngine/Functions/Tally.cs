// Keep this file CodeMaid organised and cleaned
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

        private double[]? _numericValues;

        public Tally()
        { }

        public void AddValue(Object v)
        {
            _list.Add(v);
            _numericValues = null;
        }

        public IEnumerator<object> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public double Sum() => NumericValuesInternal().Sum();


        private IEnumerable<double> NumericValuesEnumerable()
        {
            foreach (var value in _list)
            {
                if (value is string || !(value is IEnumerable vEnumerable))
                {
                    if (TryParseToDouble(value, aggressiveConversion: false, out double tmp))
                        yield return tmp;
                }
                else
                {
                    foreach (var v in vEnumerable)
                    {
                        if (TryParseToDouble(v, aggressiveConversion: false, out double tmp))
                            yield return tmp;
                    }
                }
            }
        }

        private double[] NumericValuesInternal()
                    => LazyInitializer.EnsureInitialized(ref _numericValues, () => NumericValuesEnumerable().ToArray())!;

        // If aggressiveConversion == true, then try to parse non-numeric types to double too
        private bool TryParseToDouble(object value, bool aggressiveConversion, out double d)
        {
            d = 0;
            if (value.IsNumber())
            {
                d = Convert.ToDouble(value);
                return true;
            }
            else if (value is Boolean b)
            {
                if (!aggressiveConversion) return false;

                d = (b ? 1 : 0);
                return true;
            }
            else if (value is DateTime dt)
            {
                d = dt.ToOADate();
                return true;
            }
            else if (value is TimeSpan ts)
            {
                d = ts.TotalDays;
                return true;
            }
            else if (value is string s)
            {
                if (!aggressiveConversion) return false;
                return double.TryParse(s, out d);
            }

            return false;
        }
    }
}
