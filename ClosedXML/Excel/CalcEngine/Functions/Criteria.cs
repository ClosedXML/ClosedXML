using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine.Functions;

/// <summary>
/// A representation of selection criteria used in IFs functions <c>{SUM,AVERAGE,COUNT}{IF,IFS}</c>
/// and database functions (<c>D{AVERAGE,COUNT,COUNTA,...}</c>).
/// </summary>
internal class Criteria
{
    // Values are ordered by length of a prefix. The longer ones are before sorter ones.
    private static readonly List<(string Prefix, Comparison Comparison)> AllComparisons = new()
    {
        ("<>", Comparison.NotEqual),
        (">=", Comparison.GreaterOrEqualTo),
        ("<=", Comparison.LessOrEqualTo),
        ("=", Comparison.Equal),
        (">", Comparison.GreaterThan),
        ("<", Comparison.LessThan),
    };

    private readonly Comparison _comparison;
    private readonly ScalarValue _value;
    private readonly CultureInfo _culture;

    private Criteria(Comparison comparison, ScalarValue value, CultureInfo culture)
    {
        _comparison = comparison;
        _value = value;
        _culture = culture;
    }

    /// <summary>
    /// Can a blank value match the criteria?
    /// </summary>
    internal bool CanBlankValueMatch
    {
        get
        {
            // Criteria accepts only values equal to blank (it's either blank or empty text).
            // Therefore blank values must be included, because blank is equal to blank.
            if (_comparison is Comparison.Equal or Comparison.None && _value.IsBlank)
                return true;

            // Criteria accepts only values that are not a concrete value. Blank values
            // are not concrete values and therefore must be included.
            if (_comparison == Comparison.NotEqual && !_value.IsBlank)
                return true;

            return false;
        }
    }

    internal static Criteria Create(ScalarValue criteria, CultureInfo culture)
    {
        if (criteria.IsText)
        {
            // Criteria as a text is the most common type. Text can be either comparison
            // with a value (e.g. ">7,5") or just value ("7,5"). Comparison must start
            // at the very first char, otherwise it's not interpreted as a comparison.
            var criteriaText = criteria.GetText();
            var (prefix, comparison) = GetComparison(criteriaText);
            var operandText = criteriaText[prefix.Length..];
            var operand = ScalarValue.Parse(operandText, culture);
            return new Criteria(comparison, operand, culture);
        }

        // If criteria is real blank (either through cell reference or IF(TRUE,))
        // it is interpreted as number 0.
        if (criteria.IsBlank)
            return new Criteria(Comparison.Equal, 0, culture);

        return new Criteria(Comparison.None, criteria, culture);

        static (string Prefix, Comparison Comparison) GetComparison(string criteriaText)
        {
            foreach (var (prefix, prefixComparison) in AllComparisons)
            {
                if (criteriaText.StartsWith(prefix))
                    return (prefix, prefixComparison);
            }

            return (string.Empty, Comparison.None);
        }
    }

    internal bool Match(ScalarValue value)
    {
        return _value switch
        {
            { IsBlank: true } => CompareBlank(value),
            { IsLogical: true } => CompareLogical(value, _value.GetLogical()),
            { IsNumber: true } => CompareNumber(value, _value.GetNumber()),
            { IsText: true } => CompareText(value, _value.GetText()),
            { IsError: true } => CompareError(value, _value.GetError()),
            _ => throw new UnreachableException(),
        };
    }

    private bool CompareBlank(ScalarValue value)
    {
        // This path can one be achieved when criteria was empty string (e.g. "")
        // or some comparison and empty string (e.g. "="). If the value was real
        // blank, it is interpreted as "=0"

        // Passed criteria is "". That is true only for empty string or blank
        if (_comparison == Comparison.None)
            return value.IsBlank || (value.IsText && value.GetText().Length == 0);

        // Passed criteria is "=". That is true only for blank
        if (_comparison == Comparison.Equal)
            return value.IsBlank;

        // Passed criteria is "<>". That is true only when argument is not blank.
        if (_comparison == Comparison.NotEqual)
            return !value.IsBlank;

        // Only sortable comparisons are left (>, <, >=, <=). That never makes
        // sense for blanks or other types is thus always false.
        return false;
    }

    private bool CompareLogical(ScalarValue value, bool actual)
    {
        if (!value.IsLogical)
            return _comparison == Comparison.NotEqual;

        return Compare(value.GetLogical().CompareTo(actual));
    }

    private bool CompareNumber(ScalarValue value, double actual)
    {
        double number;
        if (value.IsNumber)
        {
            number = value.GetNumber();
        }
        else if (value.IsText && ScalarValue.TextToNumber(value.GetText(), _culture).TryPickT0(out var parsedNumber, out _))
        {
            number = parsedNumber;
        }
        else
        {
            return _comparison == Comparison.NotEqual;
        }

        return Compare(number.CompareTo(actual));
    }

    private bool CompareText(ScalarValue value, string actual)
    {
        if (!value.IsText)
            return _comparison == Comparison.NotEqual;

        return _comparison switch
        {
            Comparison.Equal or Comparison.None => new Wildcard(actual).Matches(value.GetText().AsSpan()),
            Comparison.NotEqual => !new Wildcard(actual).Matches(value.GetText().AsSpan()),
            Comparison.LessThan => _culture.CompareInfo.Compare(value.GetText(), actual) < 0,
            Comparison.LessOrEqualTo => _culture.CompareInfo.Compare(value.GetText(), actual) <= 0,
            Comparison.GreaterThan => _culture.CompareInfo.Compare(value.GetText(), actual) > 0,
            Comparison.GreaterOrEqualTo => _culture.CompareInfo.Compare(value.GetText(), actual) >= 0,
            _ => throw new UnreachableException()
        };
    }

    private bool CompareError(ScalarValue value, XLError actual)
    {
        if (!value.IsError)
            return _comparison == Comparison.NotEqual;

        return Compare(value.GetError().CompareTo(actual));
    }

    private bool Compare(int cmp)
    {
        return _comparison switch
        {
            Comparison.Equal or Comparison.None => cmp == 0,
            Comparison.NotEqual => cmp != 0,
            Comparison.LessThan => cmp < 0,
            Comparison.LessOrEqualTo => cmp <= 0,
            Comparison.GreaterThan => cmp > 0,
            Comparison.GreaterOrEqualTo => cmp >= 0,
            _ => throw new UnreachableException()
        };
    }

    private enum Comparison
    {
        /// <summary>
        /// There has to be a None comparison, because criteria empty string ("")
        /// matches blank and empty string. That is not same as "=" or actual
        /// blank value. Thus it can't be reduced to equal with some operand
        /// and has to have a special case.
        /// </summary>
        None,
        Equal,
        NotEqual,
        LessThan,
        LessOrEqualTo,
        GreaterThan,
        GreaterOrEqualTo,
    }
}

