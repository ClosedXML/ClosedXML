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
        private const int BlankValue = 0;
        private const int LogicalValue = 1;
        private const int NumberValue = 2;
        private const int TextValue = 3;
        private const int ErrorValue = 4;

        private readonly byte _index;
        private readonly bool _logical;
        private readonly double _number;
        private readonly string _text;
        private readonly XLError _error;

        private ScalarValue(byte index, bool logical, double number, string text, XLError error)
        {
            _index = index;
            _logical = logical;
            _number = number;
            _text = text;
            _error = error;
        }

        /// <summary>
        /// A blank value of a scalar. It can behave as a 0 or empty string, depending on context.
        /// </summary>
        /// <example><c>A1+5</c> is a number 5, blank behaves as 0, <c>A1 &amp; "text"</c> is a "text", blank behaves as empty string.</example>
        public static readonly ScalarValue Blank = new(BlankValue, default, default, default, default);

        public bool IsBlank => _index == BlankValue;

        public static ScalarValue From(bool logical) => new(LogicalValue, logical, default, default, default);

        public static ScalarValue From(double number) => new(NumberValue, default, number, default, default);

        public static ScalarValue From(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new ScalarValue(TextValue, default, default, text, default);
        }

        public static ScalarValue From(XLError error) => new(ErrorValue, default, default, default, error);

        public static implicit operator ScalarValue(bool logical) => From(logical);

        public static implicit operator ScalarValue(double number) => From(number);

        public static implicit operator ScalarValue(string text) => From(text);

        public static implicit operator ScalarValue(XLError error) => From(error);

        public TResult Match<TResult>(Func<TResult> transformBlank, Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<XLError, TResult> transformError)
        {
            return _index switch
            {
                BlankValue => transformBlank(),
                LogicalValue => transformLogical(_logical),
                NumberValue => transformNumber(_number),
                TextValue => transformText(_text),
                ErrorValue => transformError(_error),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1>(TParam1 param, Func<TParam1, TResult> transformBlank, Func<bool, TParam1, TResult> transformLogical, Func<double, TParam1, TResult> transformNumber, Func<string, TParam1, TResult> transformText, Func<XLError, TParam1, TResult> transformError)
        {
            return _index switch
            {
                BlankValue => transformBlank(param),
                LogicalValue => transformLogical(_logical, param),
                NumberValue => transformNumber(_number, param),
                TextValue => transformText(_text, param),
                ErrorValue => transformError(_error, param),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1, TParam2>(TParam1 param1, TParam2 param2, Func<TParam1, TParam2, TResult> transformBlank, Func<bool, TParam1, TParam2, TResult> transformLogical, Func<double, TParam1, TParam2, TResult> transformNumber, Func<string, TParam1, TParam2, TResult> transformText, Func<XLError, TParam1, TParam2, TResult> transformError)
        {
            return _index switch
            {
                BlankValue => transformBlank(param1, param2),
                LogicalValue => transformLogical(_logical, param1, param2),
                NumberValue => transformNumber(_number, param1, param2),
                TextValue => transformText(_text, param1, param2),
                ErrorValue => transformError(_error, param1, param2),
                _ => throw new InvalidOperationException()
            };
        }

        public AnyValue ToAnyValue()
        {
            return _index switch
            {
                BlankValue => AnyValue.Blank,
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
        public OneOf<string, XLError> ToText(CultureInfo culture)
        {
            return _index switch
            {
                BlankValue => string.Empty,
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
        public OneOf<double, XLError> ToNumber(CultureInfo culture)
        {
            return _index switch
            {
                BlankValue => 0,
                LogicalValue => _logical ? 1.0 : 0.0,
                NumberValue => _number,
                TextValue => TextToNumber(_text, culture),
                ErrorValue => _error,
                _ => throw new InvalidOperationException()
            };
        }

        private static OneOf<double, XLError> TextToNumber(string text, CultureInfo culture)
        {
            if (string.IsNullOrWhiteSpace(text))
                return XLError.IncompatibleValue;

            // Numbers. The parsing method recognizes braces as negative number, includes currency parsing.
            // Format 1 '0'
            //        2 '0.00'
            //        3 '#,##0'
            //        4 '#,##0.00'
            //       11 '0.00E+00'
            //       48 '##0.0E+0'
            if (double.TryParse(text, NumberStyles.Any, culture, out var number))
                return number;

            // Fractions
            // Format 12 '# ?/?'
            //        13 '# ??/??'
            if (FractionParser.TryParse(text, out var fraction))
                return fraction;

            // Time span uses a different parser from time within a day.
            // Format 20 'h:mm'
            //        21 'h:mm:ss'
            //        47 'mm:ss.0'
            if (TimeSpanParser.TryParseTime(text, culture, out var timeSpan))
                return timeSpan.ToSerialDateTime();

            return XLError.IncompatibleValue;
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

        public bool TryPickError(out XLError error)
        {
            if (_index == ErrorValue)
            {
                error = _error;
                return true;
            }

            error = default;
            return false;
        }
    }
}
