using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class CalcEngineHelpers
    {
        private static Lazy<Dictionary<string, Tuple<string, string>>> patternReplacements =
            new Lazy<Dictionary<string, Tuple<string, string>>>(() =>
            {
                // key: the literal string to match
                // value: a tuple: first item: the search pattern, second item: the replacement
                return new Dictionary<string, Tuple<string, string>>()
                {
                    [@"~~"] = new Tuple<string, string>(@"~~", "~"),
                    [@"~*"] = new Tuple<string, string>(@"~\*", @"\*"),
                    [@"~?"] = new Tuple<string, string>(@"~\?", @"\?"),
                    [@"?"] = new Tuple<string, string>(@"\?", ".?"),
                    [@"*"] = new Tuple<string, string>(@"\*", ".*"),
                };
            });

        internal static bool ValueSatisfiesCriteria(object value, object criteria, CalcEngine ce)
        {
            // safety...
            if (value == null)
            {
                return false;
            }

            // Excel treats TRUE and 1 as unequal, but LibreOffice treats them as equal. We follow Excel's convention
            if (criteria is Boolean b1)
                return (value is Boolean b2) && b1.Equals(b2);

            if (value is Boolean) return false;

            // if criteria is a number, straight comparison
            Double cdbl;
            if (criteria is Double dbl2) cdbl = dbl2;
            else if (criteria is Int32 i) cdbl = i; // results of DATE function can be an integer
            else if (criteria is DateTime dt) cdbl = dt.ToOADate();
            else if (criteria is TimeSpan ts) cdbl = ts.TotalDays;
            else if (criteria is String cs)
            {
                if (value is string && (value as string).Trim().Length == 0)
                    return cs.Length == 0;

                if (cs.Length == 0)
                    return cs.Equals(value);

                // if criteria is an expression (e.g. ">20"), use calc engine
                if ((cs[0] == '=' && cs.IndexOfAny(new[] { '*', '?' }) < 0)
                    || cs[0] == '<'
                    || cs[0] == '>')
                {
                    // build expression
                    var expression = string.Format("{0}{1}", value, cs);

                    // add quotes if necessary
                    var pattern = @"([\w\s]+)(\W+)(\w+)";
                    var m = Regex.Match(expression, pattern);
                    if (m.Groups.Count == 4
                        && (!double.TryParse(m.Groups[1].Value, out double d) ||
                            !double.TryParse(m.Groups[3].Value, out d)))
                    {
                        expression = string.Format("\"{0}\"{1}\"{2}\"",
                                                   m.Groups[1].Value,
                                                   m.Groups[2].Value,
                                                   m.Groups[3].Value);
                    }

                    // evaluate
                    return (bool)ce.Evaluate(expression);
                }

                // if criteria is a regular expression, use regex
                if (cs.IndexOfAny(new[] { '*', '?' }) > -1)
                {
                    if (cs[0] == '=') cs = cs.Substring(1);

                    var pattern = Regex.Replace(
                        cs,
                        "(" + String.Join(
                                "|",
                                patternReplacements.Value.Values.Select(t => t.Item1))
                        + ")",
                        m => patternReplacements.Value[m.Value].Item2);
                    pattern = $"^{pattern}$";

                    return Regex.IsMatch(value.ToString(), pattern, RegexOptions.IgnoreCase);
                }

                // straight string comparison
                if (value is string vs)
                    return vs.Equals(cs, StringComparison.OrdinalIgnoreCase);
                else
                    return string.Equals(value.ToString(), cs, StringComparison.OrdinalIgnoreCase);
            }
            else
                throw new NotImplementedException();

            Double vdbl;
            if (value is Double dbl) vdbl = dbl;
            else if (value is Int32 i) vdbl = i;
            else if (value is DateTime dt) vdbl = dt.ToOADate();
            else if (value is TimeSpan ts) vdbl = ts.TotalDays;
            else if (value is String s)
            {
                if (!Double.TryParse(s, out vdbl)) return false;
            }
            else
                throw new NotImplementedException();

            return Math.Abs(vdbl - cdbl) < Double.Epsilon;
        }

        internal static bool ValueIsBlank(object value)
        {
            if (value == null)
                return true;

            if (value is string s)
                return s.Length == 0;

            return false;
        }

        /// <summary>
        /// Get total count of cells in the specified range without initializing them all
        /// (which might cause serious performance issues on column-wide calculations).
        /// </summary>
        /// <param name="rangeExpression">Expression referring to the cell range.</param>
        /// <returns>Total number of cells in the range.</returns>
        internal static long GetTotalCellsCount(XObjectExpression rangeExpression)
        {
            var range = (rangeExpression?.Value as CellRangeReference)?.Range;
            if (range == null)
                return 0;
            return (long)(range.LastColumn().ColumnNumber() - range.FirstColumn().ColumnNumber() + 1) *
                   (long)(range.LastRow().RowNumber() - range.FirstRow().RowNumber() + 1);
        }
    }
}
