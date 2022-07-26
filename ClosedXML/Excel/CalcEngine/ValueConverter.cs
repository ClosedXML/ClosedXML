using OneOf;
using System;
using System.Collections.Generic;
using System.Globalization;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class ValueConverter
    {
        private static readonly Dictionary<System.Type, List<System.Type>> a = new Dictionary<System.Type, List<System.Type>>()
        {
            { typeof(bool), new List<System.Type>() { typeof(double), typeof(string) } },
            { typeof(double), new List<System.Type>() { typeof(bool), typeof(string) } },
            { typeof(string), new List<System.Type>() { typeof(double) } },
            { typeof(Error), new List<System.Type>() }
        };

        private readonly CultureInfo _culture;

        public ValueConverter(CultureInfo culture) => _culture = culture;


        internal static double ToNumber(bool logical)
        {
            return logical ? 1 : 0;
        }

        internal OneOf<double, Error> ToNumber(string text)
        {
            return double.TryParse(text, NumberStyles.Float, _culture, out var number)
                ? number
                : Error.CellValue;
        }

        internal OneOf<double, Error> ToNumber(AnyValue? value)
        {
            if (!value.HasValue)
                return Error.CellValue;

            return value.Value.Match(
                    logical => ToNumber(logical),
                    number => number,
                    text => ToNumber(text),
                    error => error,
                    array => throw new NotImplementedException("Not sure what to do with it."),
                    reference => throw new NotImplementedException("Not sure what to do with it."));
        }

        internal string ToExcelString(double rightNumber)
        {
            return rightNumber.ToString(_culture);
        }

        internal OneOf<string, Error> ToText(ScalarValue lhs)
        {
            return lhs.Match<OneOf<string, Error>>(
                logical => logical ? "TRUE" : "FALSE",
                number => number.ToString(_culture),
                text => text,
                error => error);
        }

        internal OneOf<string, Error> ToText(AnyValue value)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return ToText(scalar);

            if (collection.TryPickT0(out var array, out var _))
                return ToText(array[0, 0]);

            throw new NotImplementedException();
        }
    }
}
