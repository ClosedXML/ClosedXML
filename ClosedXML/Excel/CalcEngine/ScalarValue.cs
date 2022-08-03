using System;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A representation of a value as a discriminated union.
    /// </summary>
    /// <remarks>
    /// A bare bone copy of <c>OneOf</c> that can be more optimized:
    /// <list type="bullet">
    ///   <item>readonly struct to get rid of defensive copies</item>
    ///   <item>struct can be smaler through offsets (based on NoBox)</item>
    ///   <item>allows to pass additional arguments to Match function to skip a need to instantiate a new lambda instance on each call and allow easier inlining.</item>
    /// </list>
    /// </remarks>
    internal readonly struct ScalarValue
    {
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

        public static ScalarValue From(bool logical) => new(0, logical, default, default, default);

        public static ScalarValue From(double number) => new(1, default, number, default, default);

        public static ScalarValue From(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new ScalarValue(2, default, default, text, default);
        }

        public static ScalarValue From(Error error) => new(3, default, default, default, error);

        public static implicit operator ScalarValue(bool logical) => From(logical);

        public static implicit operator ScalarValue(double number) => From(number);

        public static implicit operator ScalarValue(string text) => From(text);

        public static implicit operator ScalarValue(Error error) => From(error);

        public TResult Match<TResult>(Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<Error, TResult> transformError)
        {
            return _index switch
            {
                0 => transformLogical(_logical),
                1 => transformNumber(_number),
                2 => transformText(_text),
                3 => transformError(_error),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1>(TParam1 param, Func<bool, TParam1, TResult> transformLogical, Func<double, TParam1, TResult> transformNumber, Func<string, TParam1, TResult> transformText, Func<Error, TParam1, TResult> transformError)
        {
            return _index switch
            {
                0 => transformLogical(_logical, param),
                1 => transformNumber(_number, param),
                2 => transformText(_text, param),
                3 => transformError(_error, param),
                _ => throw new InvalidOperationException()
            };
        }

        public TResult Match<TResult, TParam1, TParam2>(TParam1 param1, TParam2 param2, Func<bool, TParam1, TParam2, TResult> transformLogical, Func<double, TParam1, TParam2, TResult> transformNumber, Func<string, TParam1, TParam2, TResult> transformText, Func<Error, TParam1, TParam2, TResult> transformError)
        {
            return _index switch
            {
                0 => transformLogical(_logical, param1, param2),
                1 => transformNumber(_number, param1, param2),
                2 => transformText(_text, param1, param2),
                3 => transformError(_error, param1, param2),
                _ => throw new InvalidOperationException()
            };
        }

        public AnyValue ToAnyValue()
        {
            return Match<AnyValue>(
                logical => logical,
                number => number,
                text => text,
                error => error);
        }
    }
}
