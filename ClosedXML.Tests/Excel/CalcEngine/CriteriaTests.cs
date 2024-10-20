using System.Collections.Generic;
using System.Globalization;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Functions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine;

[TestFixture]
internal class CriteriaTests
{
    [Test]
    [SetCulture("cs-CZ")] // cs-CZ has ',' as a decimal separator (e.g. '1,2' is one point two).
    [TestCaseSource(nameof(CriteriaTestCases))]
    public void Selection_criteria_uses_type_and_comparator_to_match_values(string selectionCriteria, XLCellValue value, bool expectedResult)
    {
        var criteria = Criteria.Create(selectionCriteria, CultureInfo.CurrentCulture);
        Assert.AreEqual(expectedResult, criteria.Match(value));
    }

    public static IEnumerable<object> CriteriaTestCases
    {
        get
        {
            // Blank without compare is interpreted as number 0
            yield return S("", 0);
            yield return S("", "0 0/2");
            yield return F("", Blank.Value);
            yield return F("", "");
            yield return F(" ", 0);
            yield return S(" ", " ");

            // Blank with equal op is interpreted as blank and type checked
            yield return S("=", Blank.Value);
            yield return F("=", 0);
            yield return F("=", "0 0/2");
            yield return F("=", "");

            // Blank with not equal is interpreted as anything but blank value
            yield return F("<>", Blank.Value);
            yield return S("<>", 0);
            yield return S("<>", "0 0/2");
            yield return S("<>", "");

            // Any comparison with blank always return false, like NaN.
            foreach (var cmp in new[] { "<", "<=", ">=", ">" })
            {
                yield return F(cmp, Blank.Value);
                yield return F(cmp, 0);
                yield return F(cmp, "0 0/2");
                yield return F(cmp, "");
            }

            // Logical are compared by type and value
            foreach (var eq in new[] { "", "=" })
            {
                yield return S(eq + "TRUE", true);
                yield return S(eq + "true", true);
                yield return F(eq + "TRUE", "TRUE");
                yield return F(eq + "TRUE", 1);
                yield return S(eq + "FALSE", false);
                yield return S(eq + "false", false);
                yield return F(eq + "FALSE", "FALSE");
                yield return F(eq + "FALSE", 0);
                yield return F(eq + "FALSE", Blank.Value);
            }

            yield return S("<>TRUE", false);
            yield return S("<>TRUE", 1);
            yield return S("<>TRUE", "Text");
            yield return S("<>TRUE", XLError.DivisionByZero);
            yield return F("<>TRUE", true);

            yield return S(">FALSE", true);
            yield return F(">FALSE", false);
            yield return F(">TRUE", true);

            yield return S(">=FALSE", true);
            yield return S(">=FALSE", false);

            yield return S("<=FALSE", false);
            yield return F("<=FALSE", true);

            yield return S("<TRUE", false);
            yield return F("<TRUE", true);

            // Number converts text, if possible
            foreach (var eq in new[] { "", "=" })
            {
                yield return S(eq + "1", 1);
                yield return S(eq + ",5", 0.5);
                yield return S(eq + "36:00", "1 1/2");
                yield return F(eq + "1,5", "text");
                yield return F(eq + "1", true);
                yield return F(eq + "1", Blank.Value);
                yield return F(eq + "1", XLError.NullValue);
            }

            yield return S("<>1", 0.9);
            yield return F("<>1", 1);
            yield return F("<>,5", "0 1/2");
            yield return S("<>1", Blank.Value);
            yield return S("<>1", true);
            yield return S("<>1", false);
            yield return S("<>1", "text");
            yield return S("<>1", XLError.NullValue);

            yield return F("<1", 1);
            yield return S("<=1", 1);
            foreach (var lt in new[] { "<", "<=" })
            {
                yield return S(lt + "1", 0);
                yield return S(lt + "1", 0.9);
                yield return S(lt + "0,5", "0,4");
                yield return S(lt + "24:00", "0 1/2");
                yield return F(lt + "24:00", "0 3/2");
                yield return F(lt + "1", "text");
                yield return F(lt + "1", "");
                yield return F(lt + "1", false);
            }

            yield return F(">1", 1);
            yield return S(">=1", 1);
            foreach (var gt in new[] { ">", ">=" })
            {
                yield return S(gt + "1", 2);
                yield return S(gt + "1", 1.1);
                yield return S(gt + "0,5", "0,6");
                yield return S(gt + "24:00", "1 1/2");
                yield return F(gt + "24:00", "0 1/2");
                yield return F(gt + "1", "text");
                yield return F(gt + "1", "");
                yield return F(gt + "0", true);
            }

            // Text for equals is a wildcard
            foreach (var eq in new[] { "", "=" })
            {
                yield return S(eq + "abc", "abc");
                yield return F(eq + "ab", "abc");
                yield return S(eq + "AbC", "aBc");
                yield return S(eq + "?", "a");
                yield return S(eq + "?", "1");
                yield return F(eq + "?", "ab");
                yield return S(eq + "a?", "ab");

                // Fail for other types
                yield return F(eq + "?", 1);
            }

            // Not equal matches with the text and then inverts the result.
            yield return F("<>?", "a");
            yield return F("<>?", "b");
            yield return S("<>?", "ab");
            yield return S("<>?", 1);
            yield return S("<>?", true);

            // Text comparison are culture dependent and don't use wildcards
            // In Czech, order of letters is 'h', 'ch', 'i' (yes, there is a two grapheme letter).
            yield return F("<a", "a");
            yield return S("<=a", "a");
            foreach (var lt in new[] { "<", "<=" })
            {
                yield return S(lt + "z", "a");
                yield return F(lt + "b", "c");
                yield return S(lt + "?", "!");
                yield return F(lt + "?", "a");
                yield return S(lt + "ch", "h"); // 'h' <= 'ch' = true
                yield return F(lt + "ch", "i"); // 'i' <= 'ch' = false.
            }

            yield return F(">a", "a");
            yield return S(">=a", "a");
            foreach (var gt in new[] { ">", ">=" })
            {
                yield return S(gt + "a", "b");
                yield return F(gt + "?", "!");
                yield return S(gt + "?", "a");
                yield return S(gt + "ch", "i"); // 'i' > 'ch' = true.
                yield return F(gt + "ch", "h"); // 'h' > 'ch' = false
            }

            // Errors
            foreach (var eq in new[] { "", "=" })
            {
                yield return S(eq + "#DIV/0!", XLError.DivisionByZero);
                yield return F(eq + "#DIV/0!", "#DIV/0!");
                yield return F(eq + "#NULL!", 1);
            }

            yield return S(">#NULL!", XLError.DivisionByZero);
            yield return F(">#DIV/0!", XLError.NullValue);

            yield return S(">=#NULL!", XLError.DivisionByZero);
            yield return S(">=#NULL!", XLError.NullValue);
            yield return F(">=#DIV/0!", XLError.NullValue);

            yield return S("<=#DIV/0!", XLError.NullValue);
            yield return S("<=#NULL!", XLError.NullValue);
            yield return F("<=#NULL!", XLError.DivisionByZero);

            yield return S("<#DIV/0!", XLError.NullValue);
            yield return F("<#NULL!", XLError.DivisionByZero);

            yield break;

            static object[] S(string s, XLCellValue v)
                => new object[] { s, v, true };

            static object[] F(string s, XLCellValue v)
                => new object[] { s, v, false };
        }
    }
}
