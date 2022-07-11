using ClosedXML.Excel.CalcEngine.Exceptions;
using ExcelNumberFormat;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Text
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ASC", 1, Asc); // Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters
            //ce.RegisterFunction("BAHTTEXT	Converts a number to text, using the ß (baht) currency format
            ce.RegisterFunction("CHAR", 1, _Char); // Returns the character specified by the code number
            ce.RegisterFunction("CLEAN", 1, Clean); //	Removes all nonprintable characters from text
            ce.RegisterFunction("CODE", 1, Code); // Returns a numeric code for the first character in a text string
            ce.RegisterFunction("CONCAT", 1, int.MaxValue, Concat); //	Joins several text items into one text item
            ce.RegisterFunction("CONCATENATE", 1, int.MaxValue, Concatenate); //	Joins several text items into one text item
            ce.RegisterFunction("DOLLAR", 1, 2, Dollar); // Converts a number to text, using the $ (dollar) currency format
            ce.RegisterFunction("EXACT", 2, Exact); // Checks to see if two text values are identical
            ce.RegisterFunction("FIND", 2, 3, Find); //Finds one text value within another (case-sensitive)
            ce.RegisterFunction("FIXED", 1, 3, Fixed); // Formats a number as text with a fixed number of decimals
            //ce.RegisterFunction("JIS	Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters
            ce.RegisterFunction("LEFT", 1, 2, Left); // LEFTB	Returns the leftmost characters from a text value
            ce.RegisterFunction("LEN", 1, Len); //, Returns the number of characters in a text string
            ce.RegisterFunction("LOWER", 1, Lower); //	Converts text to lowercase
            ce.RegisterFunction("MID", 3, Mid); // Returns a specific number of characters from a text string starting at the position you specify
            ce.RegisterFunction("NUMBERVALUE", 1, 3, NumberValue); // Converts a text argument to a number
            //ce.RegisterFunction("PHONETIC	Extracts the phonetic (furigana) characters from a text string
            ce.RegisterFunction("PROPER", 1, Proper); // Capitalizes the first letter in each word of a text value
            ce.RegisterFunction("REPLACE", 4, Replace); // Replaces characters within text
            ce.RegisterFunction("REPT", 2, Rept); // Repeats text a given number of times
            ce.RegisterFunction("RIGHT", 1, 2, Right); // Returns the rightmost characters from a text value
            ce.RegisterFunction("SEARCH", 2, 3, Search); // Finds one text value within another (not case-sensitive)
            ce.RegisterFunction("SUBSTITUTE", 3, 4, Substitute); // Substitutes new text for old text in a text string
            ce.RegisterFunction("T", 1, T); // Converts its arguments to text
            ce.RegisterFunction("TEXT", 2, _Text); // Formats a number and converts it to text
            ce.RegisterFunction("TEXTJOIN", 3, 254, TextJoin); // Joins text via delimiter
            ce.RegisterFunction("TRIM", 1, Trim); // Removes spaces from text
            ce.RegisterFunction("UPPER", 1, Upper); // Converts text to uppercase
            ce.RegisterFunction("VALUE", 1, Value); // Converts a text argument to a number
        }

        private static object _Char(List<Expression> p)
        {
            var i = (int)p[0];
            if (i < 1 || i > 255)
                throw new CellValueException(string.Format("The number {0} is out of the required range (1 to 255)", i));

            var c = (char)i;
            return c.ToString();
        }

        private static object Code(List<Expression> p)
        {
            var s = (string)p[0];
            return (int)s[0];
        }

        private static object Concat(List<Expression> p)
        {
            var sb = new StringBuilder();
            foreach (var x in p)
            {
                if (x is IEnumerable enumerable)
                {
                    foreach (var i in enumerable)
                        sb.Append((string)(new Expression(i)));
                }
                else
                    sb.Append((string)x);
            }
            return sb.ToString();
        }

        private static object Concatenate(List<Expression> p)
        {
            var sb = new StringBuilder();
            foreach (var x in p)
            {
                if (x is XObjectExpression objectExpression)
                {
                    if (objectExpression.Value is CellRangeReference cellRangeReference)
                    {
                        if (!cellRangeReference.Range.RangeAddress.IsValid)
                            throw new CellReferenceException();

                        // Only single cell range references allows at this stage. See unit test for more details
                        if (cellRangeReference.Range.RangeAddress.NumberOfCells > 1)
                            throw new CellValueException("This function does not accept cell ranges as parameters.");
                    }
                    else
                        // I'm unsure about what else objectExpression.Value could be, but let's throw CellReferenceException
                        throw new CellReferenceException();
                }

                sb.Append((string)x);
            }
            return sb.ToString();
        }

        private static object Find(List<Expression> p)
        {
            var srch = (string)p[0];
            var text = (string)p[1];
            var start = 0;
            if (p.Count > 2)
            {
                start = (int)p[2] - 1;
            }
            var index = text.IndexOf(srch, start, StringComparison.Ordinal);
            if (index == -1)
                throw new ArgumentException("String not found.");
            else
                return index + 1;
        }

        private static object Left(List<Expression> p)
        {
            var str = (string)p[0];
            var n = 1;
            if (p.Count > 1)
            {
                n = (int)p[1];
            }
            if (n >= str.Length) return str;

            return str.Substring(0, n);
        }

        private static object Len(List<Expression> p)
        {
            return ((string)p[0]).Length;
        }

        private static object Lower(List<Expression> p)
        {
            return ((string)p[0]).ToLower();
        }

        private static object Mid(List<Expression> p)
        {
            var str = (string)p[0];
            var start = (int)p[1] - 1;
            var length = (int)p[2];
            if (start > str.Length - 1)
                return String.Empty;
            if (start + length > str.Length - 1)
                return str.Substring(start);
            return str.Substring(start, length);
        }

        private static string MatchHandler(Match m)
        {
            return m.Groups[1].Value.ToUpper() + m.Groups[2].Value;
        }

        private static object Proper(List<Expression> p)
        {
            var s = (string)p[0];
            if (s.Length == 0) return "";

            MatchEvaluator evaluator = new MatchEvaluator(MatchHandler);
            StringBuilder sb = new StringBuilder();

            string pattern = "\\b(\\w)(\\w+)?\\b";
            Regex regex = new Regex(pattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            return regex.Replace(s.ToLower(), evaluator);
        }

        private static object Replace(List<Expression> p)
        {
            // old start len new
            var s = (string)p[0];
            var start = (int)p[1] - 1;
            var len = (int)p[2];
            var rep = (string)p[3];

            if (s.Length == 0) return rep;

            var sb = new StringBuilder();
            sb.Append(s.Substring(0, start));
            sb.Append(rep);
            sb.Append(s.Substring(start + len));

            return sb.ToString();
        }

        private static object Rept(List<Expression> p)
        {
            var sb = new StringBuilder();
            var s = (string)p[0];
            var repeats = (int)p[1];
            if (repeats < 0) throw new IndexOutOfRangeException("repeats");
            for (int i = 0; i < repeats; i++)
            {
                sb.Append(s);
            }
            return sb.ToString();
        }

        private static object Right(List<Expression> p)
        {
            var str = (string)p[0];
            var n = 1;
            if (p.Count > 1)
            {
                n = (int)p[1];
            }

            if (n >= str.Length) return str;

            return str.Substring(str.Length - n);
        }

        private static string WildcardToRegex(string pattern)
        {
            return Regex.Escape(pattern)
                .Replace(".", "\\.")
                .Replace("\\*", ".*")
                .Replace("\\?", ".");
        }

        private static object Search(List<Expression> p)
        {
            var search = WildcardToRegex(p[0]);
            var text = (string)p[1];

            if ("" == text) throw new ArgumentException("Invalid input string.");

            var start = 0;
            if (p.Count > 2)
            {
                start = (int)p[2] - 1;
            }

            Regex r = new Regex(search, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var match = r.Match(text.Substring(start));
            if (!match.Success)
                throw new ArgumentException("Search failed.");
            else
                return match.Index + start + 1;
            //var index = text.IndexOf(search, start, StringComparison.OrdinalIgnoreCase);
            //if (index == -1)
            //    throw new ArgumentException("String not found.");
            //else
            //    return index + 1;
        }

        private static object Substitute(List<Expression> p)
        {
            // get parameters
            var text = (string)p[0];
            var oldText = (string)p[1];
            var newText = (string)p[2];

            if ("" == text) return "";
            if ("" == oldText) return text;

            // if index not supplied, replace all
            if (p.Count == 3)
            {
                return text.Replace(oldText, newText);
            }

            // replace specific instance
            int index = (int)p[3];
            if (index < 1)
            {
                throw new ArgumentException("Invalid index in Substitute.");
            }
            int pos = text.IndexOf(oldText);
            while (pos > -1 && index > 1)
            {
                pos = text.IndexOf(oldText, pos + 1);
                index--;
            }
            return pos > -1
                ? text.Substring(0, pos) + newText + text.Substring(pos + oldText.Length)
                : text;
        }

        private static object T(List<Expression> p)
        {
            if (p[0]._token.Value?.GetType() == typeof(string))
                return (string)p[0];
            else
                return "";
        }

        private static object _Text(List<Expression> p)
        {
            var value = p[0].Evaluate();

            // Input values of type string don't get any formatting applied.
            if (value is string) return value;

            var number = (double)p[0];
            var format = (string)p[1];
            if (string.IsNullOrEmpty(format.Trim())) return "";

            var nf = new NumberFormat(format);

            if (nf.IsDateTimeFormat)
                return nf.Format(DateTime.FromOADate(number), CultureInfo.InvariantCulture);
            else
                return nf.Format(number, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// A function to Join text https://support.office.com/en-us/article/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c
        /// </summary>
        /// <param name="p">Parameters</param>
        /// <returns> string </returns>
        /// <exception cref="ApplicationException">
        /// Delimiter in first param must be a string
        /// or
        /// Second param must be a boolean (TRUE/FALSE)
        /// </exception>
        private static object TextJoin(List<Expression> p)
        {
            var values = new List<string>();
            string delimiter;
            bool ignoreEmptyStrings;
            try
            {
                delimiter = (string)p[0];
                ignoreEmptyStrings = (bool)p[1];
            }
            catch (Exception e)
            {
                throw new CellValueException("Failed to parse arguments", e);
            }

            foreach (var param in p.Skip(2))
            {
                if (param is XObjectExpression tableArray)
                {
                    if (!(tableArray.Value is CellRangeReference rangeReference))
                        throw new NoValueAvailableException("tableArray has to be a range");

                    var range = rangeReference.Range;
                    IEnumerable<string> cellValues;
                    if (ignoreEmptyStrings)
                        cellValues = range.CellsUsed()
                            .Select(c => c.GetString())
                            .Where(s => !string.IsNullOrEmpty(s));
                    else
                        cellValues = (range as XLRange).CellValues()
                            .Cast<object>()
                            .Select(o => o.ToString());

                    values.AddRange(cellValues);
                }
                else
                {
                    values.Add((string)param);
                }
            }

            var retVal = string.Join(delimiter, values);

            if (retVal.Length > 32767)
                throw new CellValueException();

            return retVal;
        }

        private static object Trim(List<Expression> p)
        {
            //Should not trim non breaking space
            //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
            return ((string)p[0]).Trim(' ');
        }

        private static object Upper(List<Expression> p)
        {
            return ((string)p[0]).ToUpper();
        }

        private static object Value(List<Expression> p)
        {
            return double.Parse(p[0], NumberStyles.Any, CultureInfo.InvariantCulture);
        }

        private static object NumberValue(List<Expression> p)
        {
            var numberFormatInfo = new NumberFormatInfo();

            numberFormatInfo.NumberDecimalSeparator = p.Count > 1 ? p[1] : CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator;
            numberFormatInfo.CurrencyDecimalSeparator = numberFormatInfo.NumberDecimalSeparator;

            numberFormatInfo.NumberGroupSeparator = p.Count > 2 ? p[2] : CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator;
            numberFormatInfo.CurrencyGroupSeparator = numberFormatInfo.NumberGroupSeparator;

            if (numberFormatInfo.NumberDecimalSeparator == numberFormatInfo.NumberGroupSeparator)
            {
                throw new CellValueException("CurrencyDecimalSeparator and CurrencyGroupSeparator have to be different.");
            }

            //Remove all whitespace characters
            var input = Regex.Replace(p[0], @"\s+", "", RegexOptions.Compiled);
            if (string.IsNullOrEmpty(input))
            {
                return 0d;
            }

            if (double.TryParse(input, NumberStyles.Any, numberFormatInfo, out var result))
            {
                if (result <= -1e308 || result >= 1e308)
                    throw new CellValueException("The value is too large");

                if (result >= -1e-309 && result <= 1e-309 && result != 0)
                    throw new CellValueException("The value is too tiny");

                if (result >= -1e-308 && result <= 1e-308)
                    result = 0d;

                return result;
            }

            throw new CellValueException("Could not convert the value to a number");
        }

        private static object Asc(List<Expression> p)
        {
            return (string)p[0];
        }

        private static object Clean(List<Expression> p)
        {
            var s = (string)p[0];

            var result = new StringBuilder();
            foreach (var c in from c in s let b = (byte)c where b >= 32 select c)
            {
                result.Append(c);
            }
            return result.ToString();
        }

        private static object Dollar(List<Expression> p)
        {
            Double value = p[0];
            int dec = p.Count == 2 ? (int)p[1] : 2;

            return value.ToString("C" + dec);
        }

        private static object Exact(List<Expression> p)
        {
            var t1 = (string)p[0];
            var t2 = (string)p[1];

            return t1 == t2;
        }

        private static object Fixed(List<Expression> p)
        {
            if (p[0]._token.Value.GetType() == typeof(string))
                throw new ApplicationException("Input type can't be string");

            Double value = p[0];
            int decimal_places = p.Count >= 2 ? (int)p[1] : 2;
            Boolean no_commas = p.Count == 3 && p[2];

            var retVal = value.ToString("N" + decimal_places);
            if (no_commas)
                return retVal.Replace(",", String.Empty);
            else
                return retVal;
        }
    }
}
