using System;
using System.Threading;

namespace ClosedXML.Excel.CalcEngine
{
    internal partial class Expression
    {
        /// <summary>
        /// Coerce a value to a different type, using Excel's convention of how date and numeric values relate
        /// </summary>
        /// <typeparam name="T">The type to which to coerce.</typeparam>
        /// <returns>The coerced value</returns>
        public T Coerce<T>()
        {
            return Coerce<T>(this._convention);
        }

        /// <summary>
        /// Coerce a value to a different type, using Excel's convention of how date and numeric values relate
        /// </summary>
        /// <typeparam name="T">The type to which to coerce.</typeparam>
        /// <param name="coercionConvention">The convention to use when converting to numbers</param>
        /// <returns>The coerced value</returns>
        public T Coerce<T>(CoercionConvention coercionConvention)
        {
            if (TryCoerce(coercionConvention, out T output))
                return output;
            else
                throw new FormatException($"Unable to coerce Expression value to type `{typeof(T).Name}`");
        }

        /// <summary>
        /// Try to coerce a value to a different type, using Excel's convention of how date and numeric values relate
        /// </summary>
        /// <typeparam name="T">The type to which to coerce.</typeparam>
        /// <param name="result">The coerced value of the new type</param>
        /// <returns>A value indicating whether coercion was successful</returns>
        public bool TryCoerce<T>(out T result)
        {
            return TryCoerce(this._convention, out result);
        }

        /// <summary>
        /// Try to coerce a value to a different type, using Excel's convention of how date and numeric values relate
        /// </summary>
        /// <typeparam name="T">The type to which to coerce.</typeparam>
        /// <param name="coercionConvention">The convention to use when converting to numbers</param>
        /// <param name="result">The coerced value of the new type</param>
        /// <returns>A value indicating whether coercion was successful</returns>
        public Boolean TryCoerce<T>(CoercionConvention coercionConvention, out T result)
        {
            if (this is ErrorExpression errorExpression)
                errorExpression.ThrowApplicableException();

            var value = this.Evaluate();

            if (value is T t)
            {
                result = t;
                return true;
            }

            var outputType = typeof(T);
            var underlyingType = outputType.GetUnderlyingType();

            if (underlyingType == typeof(string) && TryCoerceToString(value, out var s))
            {
                result = ConvertTo<T>(s);
                return true;
            }

            if (underlyingType == typeof(DateTime) && TryCoerceToDateTime(value, out var dt))
            {
                result = ConvertTo<T>(dt);
                return true;
            }

            if (underlyingType == typeof(TimeSpan) && TryCoerceToTimeSpan(value, out var ts))
            {
                result = ConvertTo<T>(ts);
                return true;
            }

            if (underlyingType == typeof(Double) && TryCoerceToDouble(value, out var dbl))
            {
                result = ConvertTo<T>(dbl);
                return true;
            }

            if (underlyingType == typeof(bool) && TryCoerceToBoolean(value, out var b))
            {
                result = ConvertTo<T>(b);
                return true;
            }

            result = default;
            return false;
        }

        private static T ConvertTo<T>(object v)
        {
            if (v is T t) return t;

            return (T)Convert.ChangeType(v, typeof(T), Thread.CurrentThread.CurrentCulture);
        }

        private static Boolean TryCoerceToDateTime(object value, out DateTime result)
        {
            if (value is TimeSpan ts)
            {
                result = new DateTime().Add(ts);
                return true;
            }

            if (value is Double dbl && dbl.IsValidOADateNumber())
            {
                result = DateTime.FromOADate(dbl);
                return true;
            }

            // handle everything else
            return TryConvertTo(value, out result);
        }

        private static Boolean TryCoerceToString(object value, out String result)
        {
            if (value == null)
            {
                result = string.Empty;
                return true;
            }

            if (value is Boolean b)
            {
                result = b.ToString().ToUpper();
                return true;
            }

            return TryConvertTo(value, out result);
        }

        private static Boolean TryCoerceToTimeSpan(object value, out TimeSpan result)
        {
            if (value is DateTime dt)
            {
                result = dt.Subtract(new DateTime());
                return true;
            }

            if (value is Double dbl)
            {
                result = XLHelper.GetTimeSpan(dbl);
                return true;
            }

            // handle everything else
            return TryConvertTo(value, out result);
        }

        private static Boolean TryConvertTo<T>(object v, out T result)
        {
            if (v is T t)
            {
                result = t;
                return true;
            }

            try
            {
                result = (T)Convert.ChangeType(v, typeof(T), Thread.CurrentThread.CurrentCulture);
                return true;
            }
            catch
            {
                result = default;
                return false;
            }
        }

        private Boolean TryCoerceToBoolean(object value, out Boolean result)
        {
            result = default;
            if (value == null)
            {
                result = false;
                return true;
            }

            // handle doubles
            if (value is Double dbl)
            {
                result = Math.Abs(dbl) > Double.Epsilon;
                return true;
            }

            if (value is string s && Boolean.TryParse(s, out var b))
            {
                result = b;
                return true;
            }

            // handle everything else
            if (TryCoerceToDouble(value, out dbl))
            {
                result = Math.Abs(dbl) > Double.Epsilon;
                return true;
            }

            return false;
        }

        private Boolean TryCoerceToDouble(object value, out Double result)
        {
            result = default;

            // handle booleans
            if (value is Boolean b)
            {
                result = b ? 1 : 0;
                return true;
            }

            // handle dates
            if (value is DateTime dt)
            {
                result = dt.ToOADate();
                return true;
            }

            if (value is TimeSpan ts)
            {
                result = ts.TotalDays;
                return true;
            }

            // handle string
            if (value is string s)
            {
                if (this._convention.HasFlag(CoercionConvention.EmptyStringAsZero) && string.IsNullOrEmpty(s)
                    || this._convention.HasFlag(CoercionConvention.NonEmptyStringAsZero) && !string.IsNullOrEmpty(s))
                {
                    result = 0;
                    return true;
                }
                else if (Double.TryParse(s, out var dbl))
                {
                    result = dbl;
                    return true;
                }
            }

            // handle nulls
            if (this._convention.HasFlag(CoercionConvention.NullAsZero) && value == null)
            {
                result = 0;
                return true;
            }

            // handle everything else
            return TryConvertTo(value, out result);
        }
    }
}
