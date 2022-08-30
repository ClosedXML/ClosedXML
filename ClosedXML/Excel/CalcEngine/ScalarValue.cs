using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A representation of a value as a discriminated union.
    /// </summary>
    /// <remarks>
    /// A bare bone copy of <c>OneOf</c> that can be more optimized:
    /// <list type="bullet">
    ///   <item>readonly struct to get rid of defensive copies</item>
    ///   <item>struct can be smaller through offsets (based on NoBox)</item>
    ///   <item>allows to pass additional arguments to Match function to skip a need to instantiate a new lambda instance on each call and allow easier inlining.</item>
    /// </list>
    /// </remarks>
    internal readonly struct ScalarValue
    {
        private const int LogicalValue = 0;
        private const int NumberValue = 1;
        private const int TextValue = 2;
        private const int ErrorValue = 3;

        private readonly byte _index;
        private readonly bool _logical;
        private readonly double _number;
        private readonly string _text;
        private readonly Error _error;

        private ScalarValue(byte index, bool logical, double number, string text, Error error)
        {
            _index = index;
            _logical = logical;
            _number = number;
            _text = text;
            _error = error;
        }

        public static ScalarValue From(bool logical) => new(LogicalValue, logical, default, default, default);

        public static ScalarValue From(double number) => new(NumberValue, default, number, default, default);

        public static ScalarValue From(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new ScalarValue(TextValue, default, default, text, default);
        }

        public static ScalarValue From(Error error) => new(ErrorValue, default, default, default, error);

        public static implicit operator ScalarValue(bool logical) => From(logical);

        public static implicit operator ScalarValue(double number) => From(number);

        public static implicit operator ScalarValue(string text) => From(text);

        public static implicit operator ScalarValue(Error error) => From(error);

        public TResult Match<TResult>(Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<Error, TResult> transformError)
        {
            return _index switch
            {
                LogicalValue => transformLogical(_logical),
                NumberValue => transformNumber(_number),
                TextValue => transformText(_text),
                ErrorValue => transformError(_error),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1>(TParam1 param, Func<bool, TParam1, TResult> transformLogical, Func<double, TParam1, TResult> transformNumber, Func<string, TParam1, TResult> transformText, Func<Error, TParam1, TResult> transformError)
        {
            return _index switch
            {
                LogicalValue => transformLogical(_logical, param),
                NumberValue => transformNumber(_number, param),
                TextValue => transformText(_text, param),
                ErrorValue => transformError(_error, param),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1, TParam2>(TParam1 param1, TParam2 param2, Func<bool, TParam1, TParam2, TResult> transformLogical, Func<double, TParam1, TParam2, TResult> transformNumber, Func<string, TParam1, TParam2, TResult> transformText, Func<Error, TParam1, TParam2, TResult> transformError)
        {
            return _index switch
            {
                LogicalValue => transformLogical(_logical, param1, param2),
                NumberValue => transformNumber(_number, param1, param2),
                TextValue => transformText(_text, param1, param2),
                ErrorValue => transformError(_error, param1, param2),
                _ => throw new InvalidOperationException()
            };
        }

        public bool TryPickNumber(out double number)
        {
            if (_index == NumberValue)
            {
                number = _number;
                return true;
            }

            number = default;
            return false;
        }

        public bool TryPickError(out Error error)
        {
            if (_index == ErrorValue)
            {
                error = _error;
                return true;
            }

            error = default;
            return false;
        }

        public AnyValue ToAnyValue()
        {
            return _index switch
            {
                LogicalValue => _logical,
                NumberValue => _number,
                TextValue => _text,
                ErrorValue => _error,
                _ => throw new InvalidOperationException()
            };
        }

        /// <summary>
        /// Convert value to text. Error is not convertible.
        /// </summary>
        public OneOf<string, Error> ToText(CultureInfo culture)
        {
            return _index switch
            {
                LogicalValue => _logical ? "TRUE" : "FALSE",
                NumberValue => _number.ToString(culture),
                TextValue => _text,
                ErrorValue => _error,
                _ => throw new InvalidOperationException()
            };
        }

        /// <summary>
        /// Convert value to number. Error is not convertible.
        /// </summary>
        public OneOf<double, Error> ToNumber(CultureInfo culture)
        {
            return _index switch
            {
                LogicalValue => _logical ? 1.0 : 0.0,
                NumberValue => _number,
                TextValue => TextToNumber(_text, culture),
                ErrorValue => _error,
                _ => throw new InvalidOperationException()
            };
        }

        private static OneOf<double, Error> TextToNumber(string text, CultureInfo culture)
        {
            return double.TryParse(text, NumberStyles.Float, culture, out var number)
                ? number
                : Error.CellValue;
        }
    }
}
