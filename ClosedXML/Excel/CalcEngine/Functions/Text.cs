using ExcelNumberFormat;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel.CalcEngine.Functions;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Text
    {
        /// <summary>
        /// Characters 0x80 to 0xFF of win-1252 encoding. Core doesn't include win-1252 encoding,
        /// so keep conversion table in this string.
        /// </summary>
        private const string Windows1252 =
            "\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F" +
            "\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178" +
            "\u00A0\u00A1\u00A2\u00A3\u00A4\u00A5\u00A6\u00A7\u00A8\u00A9\u00AA\u00AB\u00AC\u00AD\u00AE\u00AF" +
            "\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00B8\u00B9\u00BA\u00BB\u00BC\u00BD\u00BE\u00BF" +
            "\u00C0\u00C1\u00C2\u00C3\u00C4\u00C5\u00C6\u00C7\u00C8\u00C9\u00CA\u00CB\u00CC\u00CD\u00CE\u00CF" +
            "\u00D0\u00D1\u00D2\u00D3\u00D4\u00D5\u00D6\u00D7\u00D8\u00D9\u00DA\u00DB\u00DC\u00DD\u00DE\u00DF" +
            "\u00E0\u00E1\u00E2\u00E3\u00E4\u00E5\u00E6\u00E7\u00E8\u00E9\u00EA\u00EB\u00EC\u00ED\u00EE\u00EF" +
            "\u00F0\u00F1\u00F2\u00F3\u00F4\u00F5\u00F6\u00F7\u00F8\u00F9\u00FA\u00FB\u00FC\u00FD\u00FE\u00FF";

        private static readonly Lazy<Dictionary<int, string>> Windows1252Char = new(static () =>
            Enumerable.Range(0, 0x80).Select(static i => (Char: (char)i, Code: i))
                .Concat(Windows1252.Select(static (c, i) => (Char: c, Code: i + 0x80)))
                .ToDictionary(x => x.Code, x => char.ToString(x.Char)));

        private static readonly Lazy<Dictionary<char, int>> Windows1252Code = new(static () =>
            Windows1252Char.Value.ToDictionary(x => x.Value[0], x => x.Key));

        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("ASC", 1, 1, Adapt(Asc), FunctionFlags.Scalar); // Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters
            //ce.RegisterFunction("BAHTTEXT	Converts a number to text, using the ß (baht) currency format
            ce.RegisterFunction("CHAR", 1, 1, Adapt(Char), FunctionFlags.Scalar); // Returns the character specified by the code number
            ce.RegisterFunction("CLEAN", 1, 1, Adapt(Clean), FunctionFlags.Scalar); //	Removes all nonprintable characters from text
            ce.RegisterFunction("CODE", 1, 1, Adapt(Code), FunctionFlags.Scalar); // Returns a numeric code for the first character in a text string
            ce.RegisterFunction("CONCAT", 1, 255, Adapt(Concat), FunctionFlags.Future | FunctionFlags.Range, AllowRange.All); // Joins several text items into one text item
            ce.RegisterFunction("CONCATENATE", 1, 255, Adapt(Concatenate), FunctionFlags.Scalar); //	Joins several text items into one text item
            ce.RegisterFunction("DOLLAR", 1, 2, AdaptLastOptional(Dollar, 2), FunctionFlags.Scalar); // Converts a number to text, using the $ (dollar) currency format
            ce.RegisterFunction("EXACT", 2, 2, Adapt(Exact), FunctionFlags.Scalar); // Checks to see if two text values are identical
            ce.RegisterFunction("FIND", 2, 3, AdaptLastOptional(Find), FunctionFlags.Scalar); //Finds one text value within another (case-sensitive)
            ce.RegisterFunction("FIXED", 1, 3, AdaptLastTwoOptional(Fixed, 2, false), FunctionFlags.Scalar); // Formats a number as text with a fixed number of decimals
            //ce.RegisterFunction("JIS	Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters
            ce.RegisterFunction("LEFT", 1, 2, AdaptLastOptional(Left, 1), FunctionFlags.Scalar); // Returns the leftmost characters from a text value
            //ce.RegisterFunction("LEFTB", 1, 2, AdaptLastOptional(Leftb, 1), FunctionFlags.Scalar); // Returns the leftmost bytes from a text value
            ce.RegisterFunction("LEN", 1, 1, Adapt(Len), FunctionFlags.Scalar); //, Returns the number of characters in a text string
            ce.RegisterFunction("LOWER", 1, 1, Adapt(Lower), FunctionFlags.Scalar); //	Converts text to lowercase
            ce.RegisterFunction("MID", 3, 3, Adapt(Mid), FunctionFlags.Scalar); // Returns a specific number of characters from a text string starting at the position you specify
            ce.RegisterFunction("NUMBERVALUE", 1, 3, NumberValue); // Converts a text argument to a number
            //ce.RegisterFunction("PHONETIC	Extracts the phonetic (furigana) characters from a text string
            ce.RegisterFunction("PROPER", 1, Proper); // Capitalizes the first letter in each word of a text value
            ce.RegisterFunction("REPLACE", 4, Replace); // Replaces characters within text
            ce.RegisterFunction("REPT", 2, Rept); // Repeats text a given number of times
            ce.RegisterFunction("RIGHT", 1, 2, Right); // Returns the rightmost characters from a text value
            ce.RegisterFunction("SEARCH", 2, 3, AdaptLastOptional(Search), FunctionFlags.Scalar); // Finds one text value within another (not case-sensitive)
            ce.RegisterFunction("SUBSTITUTE", 3, 4, Substitute); // Substitutes new text for old text in a text string
            ce.RegisterFunction("T", 1, T); // Converts its arguments to text
            ce.RegisterFunction("TEXT", 2, _Text); // Formats a number and converts it to text
            ce.RegisterFunction("TEXTJOIN", 3, 254, TextJoin, AllowRange.Except, 0, 1); // Joins text via delimiter
            ce.RegisterFunction("TRIM", 1, Trim); // Removes spaces from text
            ce.RegisterFunction("UPPER", 1, Upper); // Converts text to uppercase
            ce.RegisterFunction("VALUE", 1, 1, Adapt(Value), FunctionFlags.Scalar); // Converts a text argument to a number
        }

        private static ScalarValue Asc(CalcContext ctx, string text)
        {
            // Excel version only works when authoring language is set to a specific languages (e.g Japanese).
            // Function doesn't do anything when Excel is set to most locales (e.g. English). There is no further
            // info. For practical purposes, it converts full-width characters from Halfwidth and Fullwidth Forms
            // unicode block to half-width variants.

            // Because fullwidth code points are in base multilingual plane, I just skip over surrogates.
            var sb = new StringBuilder(text.Length);
            foreach (int c in text)
                sb.Append((char)ToHalfForm(c));

            return sb.ToString();

            // Per ODS specification https://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part2.html#ASC
            static int ToHalfForm(int c)
            {
                return c switch
                {
                    >= 0x30A1 and <= 0x30AA when c % 2 == 0 => (c - 0x30A2) / 2 + 0xFF71, // katakana a-o
                    >= 0x30A1 and <= 0x30AA when c % 2 == 1 => (c - 0x30A1) / 2 + 0xFF67, // katakana small a-o
                    >= 0x30AB and <= 0x30C2 when c % 2 == 1 => (c - 0x30AB) / 2 + 0xFF76, // katakana ka-chi
                    >= 0x30AB and <= 0x30C2 when c % 2 == 0 => (c - 0x30AC) / 2 + 0xFF76, // katakana ga-dhi
                    0x30C3 => 0xFF6F, // katakana small tsu
                    >= 0x30C4 and <= 0x30C9 when c % 2 == 0 => (c - 0x30C4) / 2 + 0xFF82, // katakana tsu-to
                    >= 0x30C4 and <= 0x30C9 when c % 2 == 1 => (c - 0x30C5) / 2 + 0xFF82, // katakana du-do
                    >= 0x30CA and <= 0x30CE => c - 0x30CA + 0xFF85, // katakana na-no
                    >= 0x30CF and <= 0x30DD when c % 3 == 0 => (c - 0x30CF) / 3 + 0xFF8A, // katakana ha-ho
                    >= 0x30CF and <= 0x30DD when c % 3 == 1 => (c - 0x30D0) / 3 + 0xFF8A, // katakana ba-bo
                    >= 0x30CF and <= 0x30DD when c % 3 == 2 => (c - 0x30d1) / 3 + 0xff8a, // katakana pa-po
                    >= 0x30DE and <= 0x30E2 => c - 0x30DE + 0xFF8F, // katakana ma-mo
                    >= 0x30E3 and <= 0x30E8 when c % 2 == 0 => (c - 0x30E4) / 2 + 0xFF94, // katakana ya-yo
                    >= 0x30E3 and <= 0x30E8 when c % 2 == 1 => (c - 0x30E3) / 2 + 0xFF6C, // katakana small ya - yo
                    >= 0x30E9 and <= 0x30ED => c - 0x30e9 + 0xff97, // katakana ra-ro
                    0x30EF => 0xFF9C, // katakana wa
                    0x30F2 => 0xFF66, // katakana wo
                    0x30F3 => 0xFF9D, // katakana n
                    >= 0xFF01 and <= 0xFF5E => c - 0xFF01 + 0x0021, // ASCII characters
                    0x2015 => 0xFF70, // HORIZONTAL BAR => HALFWIDTH KATAKANA-HIRAGANA PROLONGED SOUND MARK
                    0x2018 => 0x0060, // LEFT SINGLE QUOTATION MARK => GRAVE ACCENT
                    0x2019 => 0x0027, // RIGHT SINGLE QUOTATION MARK => APOSTROPHE
                    0x201D => 0x0022, // RIGHT DOUBLE QUOTATION MARK => QUOTATION MARK
                    0x3001 => 0xFF64, // IDEOGRAPHIC COMMA
                    0x3002 => 0xFF61, // IDEOGRAPHIC FULL STOP
                    0x300C => 0xFF62, // LEFT CORNER BRACKET
                    0x300D => 0xFF63, // RIGHT CORNER BRACKET
                    0x309B => 0xFF9E, // KATAKANA-HIRAGANA VOICED SOUND MARK
                    0x309C => 0xFF9F, // KATAKANA-HIRAGANA SEMI-VOICED SOUND MARK
                    0x30FB => 0xFF65, // KATAKANA MIDDLE DOT
                    0x30FC => 0xFF70, // KATAKANA-HIRAGANA PROLONGED SOUND MARK
                    0xFFE5 => 0x005C, // FULLWIDTH YEN SIGN => REVERSE SOLIDUS "\"
                    _ => c
                };
            }
        }

        private static ScalarValue Char(double number)
        {
            number = Math.Truncate(number);
            if (number is < 1 or > 255)
                return XLError.IncompatibleValue;

            // Spec says to interpret numbers as values encoded in iso-8859-1. The actual
            // encoding depends on authoring language, e.g. JP uses JIS X 0201. Fun fact,
            // JP has values 253-255 from iso-8859-1, not JIS. EN/CZ/RU uses win-1252.
            // Anyway, there is no way to get a map of all encodings, so let's use one.
            // Win-1252 is probably the best default choice, because this function is
            // pre-unicode and Excel was mostly sold in US/EU.
            var value = checked((int)number);

            return Windows1252Char.Value[value];
        }

        private static ScalarValue Clean(CalcContext ctx, string text)
        {
            // Although standard says it removes only 0..1F, real one removes other characters as
            // well. Based on `LEN(CLEAN(UNICHAR(A1))) = 0`, it removes 1-1F and 0x80-0x9F. ODF
            // says to remove Cc and Cn, but Excel doesn't seem to remove Cn.
            var result = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                int codePoint = c;
                if (codePoint is >= 0 and <= 0x1F)
                    continue;

                if (codePoint is >= 0x80 and <= 0x9F)
                    continue;

                result.Append(c);
            }

            return result.ToString();
        }

        private static ScalarValue Code(CalcContext ctx, string text)
        {
            // CODE should be an inverse function to CHAR
            if (text.Length == 0)
                return XLError.IncompatibleValue;

            if (!Windows1252Code.Value.TryGetValue(text[0], out var code))
                return Windows1252Code.Value['?'];

            return code;
        }

        private static ScalarValue Concat(CalcContext ctx, List<Array> texts)
        {
            var sb = new StringBuilder();
            foreach (var array in texts)
            {
                foreach (var scalar in array)
                {
                    if (!scalar.ToText(ctx.Culture).TryPickT0(out var text, out var error))
                        return error;

                    sb.Append(text);
                    if (sb.Length > 32767)
                        return XLError.IncompatibleValue;
                }
            }

            return sb.ToString();
        }

        private static ScalarValue Concatenate(CalcContext ctx, List<string> texts)
        {
            var totalLength = texts.Sum(static x => x.Length);
            var sb = new StringBuilder(totalLength);
            foreach (var text in texts)
            {
                sb.Append(text);
                if (sb.Length > 32767)
                    return XLError.IncompatibleValue;
            }

            return sb.ToString();
        }

        private static AnyValue Find(CalcContext ctx, String findText, String withinText, OneOf<double, Blank> startNum)
        {
            var startIndex = startNum.TryPickT0(out var startNumber, out _) ? (int)Math.Truncate(startNumber) - 1 : 0;
            if (startIndex < 0 || startIndex > withinText.Length)
                return XLError.IncompatibleValue;

            var text = withinText.AsSpan(startIndex);
            var index = text.IndexOf(findText.AsSpan());
            return index == -1
                ? XLError.IncompatibleValue
                : index + startIndex + 1;
        }

        private static ScalarValue Fixed(CalcContext ctx, double number, double numDecimals, bool suppressComma)
        {
            numDecimals = Math.Truncate(numDecimals);

            // Excel allows up to 127 decimal digits. The .NET Core 8+ allows it, but older Core and
            // Fx are more limited. To keep code sane, use 99, so N99 formatting string works everywhere.
            if (numDecimals > 99)
                return XLError.IncompatibleValue;

            var culture = ctx.Culture;
            if (suppressComma)
            {
                culture = (CultureInfo)culture.Clone();
                culture.NumberFormat.NumberGroupSeparator = string.Empty;
            }

            var rounded = XLMath.Round(number, numDecimals);

            // Number rounded to tens, hundreds... should be displayed without any decimal places
            var digits = Math.Max(numDecimals, 0);
            return rounded.ToString("N" + digits, culture);
        }

        private static ScalarValue Left(CalcContext ctx, string text, double numChars)
        {
            numChars = Math.Truncate(numChars);
            if (numChars < 0)
                return XLError.IncompatibleValue;

            if (numChars >= text.Length)
                return text;

            // StringInfo.LengthInTextElements returns a length in graphemes, regardless of
            // how is grapheme stored (e.g. denormalized family emoji is 7 code points long,
            // with 4 emoji and 3 zero width joiners).
            // Generally we should return number of codepoints, at least that's how Excel and
            // LibreOffice do it (at least for LEFT).
            var i = 0;
            while (numChars > 0 && i < text.Length)
            {
                // Most C# text API will happily ignore invalid surrogate pairs, so do we
                i += char.IsSurrogatePair(text, i) ? 2 : 1;
                numChars--;
            }

            return text[..i];
        }

        private static ScalarValue Len(CalcContext ctx, string text)
        {
            // Excel counts code units, not codepoints, e.g. it returns 2 for emoji in astral
            // plane. LibreOffice returns 1 and most other functions (e.g. LEFT) use codepoints,
            // not code units. Sanity says count codepoints, but compatibility says code units.
            return text.Length;
        }

        private static ScalarValue Lower(CalcContext ctx, string text)
        {
            // Spec says "by doing a character-by-character conversion"
            // so don't do the whole string at once.
            var sb = new StringBuilder(text.Length);
            for (var i = 0; i < text.Length; ++i)
            {
                var c = text[i];
                char lowercase;
                if (i == text.Length - 1 && c == 'Σ')
                {
                    // Spec: when Σ (U+03A3) is found in a word-final position, it is converted
                    // to ς (U+03C2) instead of σ (U+03C3).
                    lowercase = 'ς';
                }
                else
                {
                    lowercase = char.ToLower(c, ctx.Culture);
                }

                sb.Append(lowercase);
            }

            return sb.ToString();
        }

        private static ScalarValue Mid(CalcContext ctx, string text, double startPos, double numChars)
        {
            // Unlike LEFT, MID uses code units and even cuts off half of surrogates,
            // e.g. LEN(MID("😊😊",1,3)) = 3. Also, spec has parameters at wrong places.
            if (startPos is < 1 or >= int.MaxValue + 1d || numChars is < 0 or >= int.MaxValue + 1d)
                return XLError.IncompatibleValue;

            var start = checked((int)Math.Truncate(startPos)) - 1;
            var length = checked((int)Math.Truncate(numChars));
            if (start >= text.Length - 1)
                return string.Empty;

            if (start + length >= text.Length)
                return text[start..];

            return text.Substring(start, length);
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

        private static AnyValue Search(CalcContext ctx, String findText, String withinText, OneOf<double, Blank> startNum)
        {
            if (withinText.Length == 0)
                return XLError.IncompatibleValue;

            var startIndex = startNum.TryPickT0(out var startNumber, out _) ? (int)Math.Truncate(startNumber) : 1;
            startIndex -= 1;
            if (startIndex < 0 || startIndex >= withinText.Length)
                return XLError.IncompatibleValue;

            var wildcard = new Wildcard(findText);
            ReadOnlySpan<char> text = withinText.AsSpan().Slice(startIndex);
            var firstIdx = wildcard.Search(text);
            if (firstIdx < 0)
                return XLError.IncompatibleValue;

            return firstIdx + startIndex + 1;
        }

        private static object Substitute(List<Expression> p)
        {
            // get parameters
            var text = (string)p[0];
            var oldText = (string)p[1];
            var newText = (string)p[2];

            if (text.Length == 0) return "";
            if (oldText.Length == 0) return text;

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
            var value = p[0].Evaluate();
            if (value is string)
                return value;
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
            catch (Exception)
            {
                return XLError.IncompatibleValue;
            }

            foreach (var param in p.Skip(2))
            {
                if (param is XObjectExpression tableArray)
                {
                    if (!(tableArray.Value is CellRangeReference rangeReference))
                        return XLError.NoValueAvailable;

                    var range = rangeReference.Range;
                    IEnumerable<string> cellValues;
                    if (ignoreEmptyStrings)
                        cellValues = range.CellsUsed()
                            .Select(c => c.GetString())
                            .Where(s => !string.IsNullOrEmpty(s));
                    else
                        cellValues = rangeReference.CellValues()
                            .Select(o => o.ToString(CultureInfo.CurrentCulture));

                    values.AddRange(cellValues);
                }
                else
                {
                    values.Add((string)param);
                }
            }

            var retVal = string.Join(delimiter, values);

            if (retVal.Length > 32767)
                return XLError.IncompatibleValue;

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

        private static AnyValue Value(CalcContext ctx, ScalarValue arg)
        {
            // Specification is vague/misleading:
            // * function accepts significantly more diverse range of inputs e.g. result of "($100)" is -100
            //   despite braces not being part of any default number format.
            // * Different cultures work weird, e.g. 7:30 PM is detected as 19:30 in cs locale despite "PM" designator being "odp."
            // * Formats 14 and 22 differ depending on the locale (that is why in dialogue are with a '*' sign)
            if (arg.IsBlank)
                return 0;

            if (arg.TryPickNumber(out var number))
                return number;

            if (!arg.TryPickText(out var text, out var error))
                return error;

            const string percentSign = "%";
            var isPercent = text.IndexOf(percentSign, StringComparison.Ordinal) >= 0;
            var textWithoutPercent = isPercent ? text.Replace(percentSign, string.Empty) : text;
            if (double.TryParse(textWithoutPercent, NumberStyles.Any, ctx.Culture, out var parsedNumber))
                return isPercent ? parsedNumber / 100d : parsedNumber;

            // fraction not parsed, maybe in the future
            // No idea how Date/Time parsing works, good enough for initial approach
            var dateTimeFormats = new[]
            {
                ctx.Culture.DateTimeFormat.ShortDatePattern,
                ctx.Culture.DateTimeFormat.YearMonthPattern,
                ctx.Culture.DateTimeFormat.ShortTimePattern,
                ctx.Culture.DateTimeFormat.LongTimePattern,
                @"mm-dd-yy", // format 14
                @"d-MMMM-yy", // format 15
                @"d-MMMM", // format 16
                @"d-MMM-yyyy",
                @"H:mm", // format 20
                @"H:mm:ss" // format 21
            };
            const DateTimeStyles dateTimeStyle = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;
            if (DateTime.TryParseExact(text, dateTimeFormats, ctx.Culture, dateTimeStyle, out var parsedDate))
                return parsedDate.ToOADate();

            return XLError.IncompatibleValue;
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
                return XLError.IncompatibleValue;
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
                    return XLError.IncompatibleValue;

                if (result >= -1e-309 && result <= 1e-309 && result != 0)
                    return XLError.IncompatibleValue;

                if (result >= -1e-308 && result <= 1e-308)
                    result = 0d;

                return result;
            }

            return XLError.IncompatibleValue;
        }

        private static ScalarValue Dollar(CalcContext ctx, double number, double decimals)
        {
            // Excel has limit of 127 decimal places, but C# has limit of 99.
            decimals = Math.Truncate(decimals);
            if (decimals > 99)
                return XLError.IncompatibleValue;

            if (decimals >= 0)
                return number.ToString("C" + decimals, ctx.Culture);

            var factor = Math.Pow(10, -decimals);
            var rounded = Math.Round(number / factor, 0, MidpointRounding.AwayFromZero);
            if (rounded != 0)
                rounded *= factor;

            return rounded.ToString("C0", ctx.Culture);
        }

        private static ScalarValue Exact(string lhs, string rhs)
        {
            return lhs == rhs;
        }
    }
}
