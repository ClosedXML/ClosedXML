﻿using OneOf;
using System.Collections.Generic;
using System.Globalization;

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
    }
}
