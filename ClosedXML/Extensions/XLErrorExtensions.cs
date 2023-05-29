using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ClosedXML.Extensions
{
    internal static class XLErrorExtensions
    {
        public static string ToDisplayString(this XLError error) =>
            error switch
            {
                XLError.CellReference => "#REF!",
                XLError.IncompatibleValue => "#VALUE!",
                XLError.DivisionByZero => "#DIV/0!",
                XLError.NameNotRecognized => "#NAME?",
                XLError.NoValueAvailable => "#N/A",
                XLError.NullValue => "#NULL!",
                XLError.NumberInvalid => "#NUM!",
                _ => throw new ArgumentOutOfRangeException()
            };
    }

    internal static class XLErrorParser
    {
        private static readonly Dictionary<string, XLError> ErrorMap = new(StringComparer.Ordinal)
        {
            ["#REF!"] = XLError.CellReference,
            ["#VALUE!"] = XLError.IncompatibleValue,
            ["#DIV/0!"] = XLError.DivisionByZero,
            ["#NAME?"] = XLError.NameNotRecognized,
            ["#N/A"] = XLError.NoValueAvailable,
            ["#NULL!"] = XLError.NullValue,
            ["#NUM!"] = XLError.NumberInvalid
        };

        public static bool TryParseError(String input, out XLError error)
            => ErrorMap.TryGetValue(input.Trim(), out error);
    }
}
