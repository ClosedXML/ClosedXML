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

            // Percents. Percent sign can be at both sides.
            // Format 9 '0%'
            //       10 '0.00%'
            var textSpan = text.AsSpan(); // Avoid extra allocations for trimming/substrings if not match
            var textSpanTrimmedEnd = textSpan.TrimEnd();
            var percentSymbol = culture.NumberFormat.PercentSymbol.AsSpan();
            if (textSpanTrimmedEnd.EndsWith(percentSymbol))
                return ParsePercent(text, 0, textSpanTrimmedEnd.Length - percentSymbol.Length, culture);

            var textSpanTrimmedStart = textSpan.TrimStart();
            if (textSpanTrimmedStart.StartsWith(percentSymbol))
            {
                var newStart = text.Length - textSpanTrimmedStart.Length + percentSymbol.Length;
                return ParsePercent(text, newStart, text.Length - newStart, culture);
            }

            // Fractions
            // Format 12 '# ?/?'
            //        13 '# ??/??'
            if (FractionParser.TryParse(text, out var fraction))
                return fraction;

            const DateTimeStyles dateStyle = DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowInnerWhite | DateTimeStyles.AllowTrailingWhite;

            // This date varies by the culture. Keep first before other standard patterns. Must be for both yy and yyyy.
            // Format 14 : short date (for en 'm/d/yyyy')
            // Format 22 : short date + hours (for en 'm/d/yyyy h:mm')
            if (DateTimeParser.TryParseCultureDate(text, culture, out var dateFormat14Or22))
                return ToSerialDate(dateFormat14Or22);

            // Date with names of months. The names of months differ across cultures.
            // Format 15 'd-mmm-yy'
            if (DateTime.TryParseExact(text, new[] { "d-MMM-yyyy", "d-MMMM-yyyy", "d-MMM-yy", "d-MMMM-yy" }, culture, dateStyle, out var dateFormat15))
                return ToSerialDate(dateFormat15);

            // Since format doesn't have a year, it uses current year 
            // Format 16 'd-mmm'
            if (DateTime.TryParseExact(text, new[] { "d-MMM", "d-MMMM" }, culture, dateStyle, out var dateFormat16))
                return ToSerialDate(dateFormat16);

            // Month and a number. In some cultures, the culture date parsing will interpret this pattern as MMM-dd, but
            // that depends on culture date patterns above. Use MMM and MMMM to encompass both abbreviation and full name.
            // Format 17 'mmm-yy'
            if (DateTime.TryParseExact(text, new[] { "MMM-y", "MMMM-y" }, culture, dateStyle, out var dateFormat17))
            {
                if (dateFormat17.Year != DateTime.Now.Year && dateFormat17.Year >= 2030)
                    dateFormat17 = dateFormat17.AddYears(-100);

                return ToSerialDate(dateFormat17);
            }

            // Format 18 'h:mm AM/PM', works for both localized and AM/PM literal
            // Format 19 'h:mm:ss AM/PM'
            if (DateTimeParser.TryParseTimeOfDay(text, culture, out var dateFormat18Or19))
                return dateFormat18Or19.ToOADate();

            // Time span uses a different parser from time of a day.
            // Format 20 'h:mm'
            //        21 'h:mm:ss'
            //        47 'mm:ss.0'
            if (TimeSpanParser.TryParseTime(text, culture, out var timeSpan))
                return timeSpan.ToSerialDateTime();

            return XLError.IncompatibleValue;

            static OneOf<double, XLError> ParsePercent(string text, int start, int length, CultureInfo c)
            {
                text = text.Substring(start, length);
                if (double.TryParse(text, NumberStyles.Float
                                          | NumberStyles.AllowThousands
                                          | NumberStyles.AllowParentheses, c, out var percents))
                    return percents / 100;

                // other formats don't use '%' sign, but text has it, so just stop for invalid inputs like 'hundred%'
                return XLError.IncompatibleValue;
            }

            static OneOf<double, XLError> ToSerialDate(DateTime dateTime)
            {
                if (dateTime.Year < 1900)
                    return XLError.IncompatibleValue;

                // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
                // to Lotus 1-2-3 decision from 1983...
                var oDate = dateTime.ToOADate();
                const int nonExistent1900Feb29SerialDate = 60;
                return oDate <= nonExistent1900Feb29SerialDate ? oDate - 1 : oDate;
            }
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

        public bool TryPickText(out string text, out XLError error)
        {
            if (_index == TextValue)
            {
                text = _text;
                error = default;
                return true;
            }

            text = default;
            error = _index == ErrorValue ? _error : XLError.IncompatibleValue;
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
