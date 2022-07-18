using OneOf;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class ValueConverter
    {
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
    }
}
