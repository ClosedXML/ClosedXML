using OneOf;
using System;
using System.Collections.Generic;
using System.Globalization;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class ValueConverter
    {
        private static readonly Dictionary<System.Type, List<System.Type>> a = new Dictionary<System.Type, List<System.Type>>()
        {
            { typeof(Logical), new List<System.Type>() { typeof(Number1), typeof(Text) } },
            { typeof(Number1), new List<System.Type>() { typeof(Logical), typeof(Text) } },
            { typeof(Text), new List<System.Type>() { typeof(Number1) } },
            { typeof(Error1), new List<System.Type>() }
        };

        private readonly CultureInfo _culture;

        public ValueConverter(CultureInfo culture) => _culture = culture;


        internal Number1 ToNumber(Logical logical)
        {
            return logical ? Number1.One : Number1.Zero;
        }

        internal OneOf<Number1, Error1> ToNumber(Text text)
        {
            return double.TryParse(text.Value, NumberStyles.Float, _culture, out var number)
                ? new Number1(number)
                : Error1.Value;
        }

        internal OneOf<Number1, Error1> ToNumber(AnyValue? value)
        {
            if (!value.HasValue)
                return Error1.Value;

            return value.Value.Match(
                    logical => ToNumber(logical),
                    number => number,
                    text => ToNumber(text),
                    error => error,
                    array => throw new NotImplementedException("Not sure what to do with it."),
                    reference => throw new NotImplementedException("Not sure what to do with it."));
        }

        internal string ToExcelString(Number1 rightNumber)
        {
            return rightNumber.Value.ToString(_culture);
        }

        internal OneOf<Text, Error1> ToText(ScalarValue lhs)
        {
            return lhs.Match<OneOf<Text, Error1>>(
                logical => new Text(logical ? "TRUE" : "FALSE"),
                number => new Text(number.Value.ToString(_culture)),
                text => text,
                error => error);
        }
    }
}
