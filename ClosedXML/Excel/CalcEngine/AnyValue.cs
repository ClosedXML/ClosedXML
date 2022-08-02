using System;
using CollectionValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    internal readonly struct AnyValue
    {
        private readonly byte _index;
        private readonly bool _logical;
        private readonly double _number;
        private readonly string _text;
        private readonly Error _error;
        private readonly Array _array;
        private readonly Reference _reference;

        private AnyValue(byte index, bool logical, double number, string text, Error error, Array array, Reference reference)
        {
            _index = index;
            _logical = logical;
            _number = number;
            _text = text;
            _error = error;
            _array = array;
            _reference = reference;
        }

        public static AnyValue FromT0(bool logical) => new(0, logical, default, default, default, default, default);

        public static AnyValue FromT1(double number) => new(1, default, number, default, default, default, default);

        public static AnyValue FromT2(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new AnyValue(2, default, default, text, default, default, default);
        }

        public static AnyValue FromT3(Error error) => new(3, default, default, default, error, default, default);

        public static AnyValue FromT4(Array array)
        {
            if (array is null)
                throw new ArgumentNullException();

            return new(4, default, default, default, default, array, default);
        }

        public static AnyValue FromT5(Reference reference)
        {
            if (reference is null)
                throw new ArgumentNullException();

            return new(5, default, default, default, default, default, reference);
        }

        public static implicit operator AnyValue(bool logical) => FromT0(logical);

        public static implicit operator AnyValue(double number) => FromT1(number);

        public static implicit operator AnyValue(string text) => FromT2(text);

        public static implicit operator AnyValue(Error error) => FromT3(error);

        public static implicit operator AnyValue(Array array) => FromT4(array);

        public static implicit operator AnyValue(Reference reference) => FromT5(reference);

        public bool TryPickScalar(out ScalarValue scalar, out CollectionValue collection)
        {
            scalar = _index switch
            {
                0 => _logical,
                1 => _number,
                2 => _text,
                3 => _error,
                _ => default
            };
            collection = _index switch
            {
                4 => _array,
                5 => _reference,
                _ => default
            };
            return _index <= 3;
        }

        public bool TryPickReference(out Reference reference)
        {
            if (_index == 5)
            {
                reference = _reference;
                return true;
            }

            reference = default;
            return false;
        }

        public TResult Match<TResult>(Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<Error, TResult> transformError, Func<Array, TResult> transformArray, Func<Reference, TResult> transformReference)
        {
            return _index switch
            {
                0 => transformLogical(_logical),
                1 => transformNumber(_number),
                2 => transformText(_text),
                3 => transformError(_error),
                4 => transformArray(_array),
                5 => transformReference(_reference),
                _ => throw new InvalidOperationException()
            };
        }
    }
}
