using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class ValueConverter
    {
        private readonly CultureInfo _culture;
        private readonly CalcContext _ctx;

        public ValueConverter(CultureInfo culture, CalcContext ctx)
        {
            _culture = culture;
            _ctx = ctx;
        }

        internal OneOf<double, Error> ToNumber(string text)
        {
            return TextToNumber(_culture, text);
        }

        private static  OneOf<double, Error> TextToNumber(CultureInfo culture, string text)
        {
            return double.TryParse(text, NumberStyles.Float, culture, out var number)
                ? number
                : Error.CellValue;
        }


        public OneOf<double, Error> CovertToNumber(ScalarValue value)
        {
            return value.Match(this,
                (logical, _) => logical ? 1 : 0,
                (number, _) => number,
                (text, conv) => conv.ToNumber(text),
                (error, _) => error);
        }

        internal OneOf<double, Error> ToNumber(AnyValue? value)
        {
            if (!value.HasValue)
                return Error.CellValue;

            if (value.Value.TryPickScalar(out var scalar, out var collection))
                return ScalarToNumber(scalar, _culture);

            return collection.Match(
                    array => throw new NotImplementedException("Not sure what to do with it."),
                    reference =>
                    {
                        if (reference.TryGetSingleCellValue(out var scalarValue, _ctx))
                            return ScalarToNumber(scalarValue, _culture);

                        throw new NotImplementedException("Not sure what to do with it.");
                    });

            static OneOf<double, Error> ScalarToNumber(ScalarValue value, CultureInfo culture)
            {
                return value.Match(culture,
                        (logical, _) => logical ? 1.0 : 0.0,
                        (number, _) => number,
                        (text, culture) => TextToNumber(culture, text),
                        (error, _) => error);
            }
        }

        internal string ToExcelString(double rightNumber)
        {
            return rightNumber.ToString(_culture);
        }

        internal OneOf<string, Error> ToText(ScalarValue lhs)
        {
            return lhs.Match<OneOf<string, Error>, CultureInfo>(_culture,
                (logical, _) => logical ? "TRUE" : "FALSE",
                (number, culture) => number.ToString(culture),
                (text, _) => text,
                (error, _) => error);
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
