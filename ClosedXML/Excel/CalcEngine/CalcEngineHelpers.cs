using System;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcEngineHelpers
    {
        internal static bool ValueSatisfiesCriteria(object value, object criteria, CalcEngine ce)
        {
            // safety...
            if (value == null)
            {
                return false;
            }

            // if criteria is a number, straight comparison
            if (criteria is double)
            {
                if (value is Double)
                    return (double)value == (double)criteria;
                Double dValue;
                return Double.TryParse(value.ToString(), out dValue) && dValue == (double)criteria;
            }

            // convert criteria to string
            var cs = criteria as string;
            if (!string.IsNullOrEmpty(cs))
            {
                // if criteria is an expression (e.g. ">20"), use calc engine
                if (cs[0] == '=' || cs[0] == '<' || cs[0] == '>')
                {
                    // build expression
                    var expression = string.Format("{0}{1}", value, cs);

                    // add quotes if necessary
                    var pattern = @"(\w+)(\W+)(\w+)";
                    var m = Regex.Match(expression, pattern);
                    if (m.Groups.Count == 4)
                    {
                        double d;
                        if (!double.TryParse(m.Groups[1].Value, out d) ||
                            !double.TryParse(m.Groups[3].Value, out d))
                        {
                            expression = string.Format("\"{0}\"{1}\"{2}\"",
                                                       m.Groups[1].Value,
                                                       m.Groups[2].Value,
                                                       m.Groups[3].Value);
                        }
                    }

                    // evaluate
                    return (bool)ce.Evaluate(expression);
                }

                // if criteria is a regular expression, use regex
                if (cs.IndexOf('*') > -1)
                {
                    var pattern = cs.Replace(@"\", @"\\");
                    pattern = pattern.Replace(".", @"\");
                    pattern = pattern.Replace("*", ".*");
                    return Regex.IsMatch(value.ToString(), pattern, RegexOptions.IgnoreCase);
                }

                // straight string comparison
                return string.Equals(value.ToString(), cs, StringComparison.OrdinalIgnoreCase);
            }

            // should never get here?
            Debug.Assert(false, "failed to evaluate criteria in SumIf");
            return false;
        }
    }
}
