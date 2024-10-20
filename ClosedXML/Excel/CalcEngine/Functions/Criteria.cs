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

    internal static Criteria Create(string text, CultureInfo culture)
    {
        // There can't be space at the start, comparison must start at first char
        var comparison = Comparison.Equal;
        var prefixLength = 0;
        foreach (var (prefix, prefixComparison) in AllComparisons)
        {
            if (text.StartsWith(prefix))
            {
                comparison = prefixComparison;
                prefixLength = prefix.Length;
                break;
            }
        }

        var value = XLCellValue.FromText(text[prefixLength..], culture);

        if (value.IsBlank)
        {
            // Empty string is matched as number 0
            return text.Length > 0
                ? new Criteria(comparison, ScalarValue.Blank, culture)
                : new Criteria(comparison, 0, culture);
        }

        if (value.IsBoolean)
            return new Criteria(comparison, value.GetBoolean(), culture);

        if (value.IsUnifiedNumber)
            return new Criteria(comparison, value.GetUnifiedNumber(), culture);

        if (value.IsText)
            return new Criteria(comparison, value.GetText(), culture);

        return new Criteria(comparison, value.GetError(), culture);
    }

    internal bool Match(XLCellValue value)
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

    private bool CompareBlank(XLCellValue value)
    {
        if (!value.IsBlank)
            return _comparison == Comparison.NotEqual;

        // Any comparison with a blank doesn't make sense and always returns false.
        // Both values are blank, so only equal matches.
        return _comparison == Comparison.Equal;
    }

    private bool CompareLogical(XLCellValue value, bool actual)
    {
        if (!value.IsBoolean)
            return _comparison == Comparison.NotEqual;

        return Compare(value.GetBoolean().CompareTo(actual));
    }

    private bool CompareNumber(XLCellValue value, double actual)
    {
        double number;
        if (value.IsUnifiedNumber)
        {
            number = value.GetUnifiedNumber();
        }
        else if (value.TryGetText(out var text) &&
                 ScalarValue.TextToNumber(text, _culture).TryPickT0(out var parsedNumber, out _))
        {
            number = parsedNumber;
        }
        else
        {
            return _comparison == Comparison.NotEqual;
        }

        return Compare(number.CompareTo(actual));
    }

    private bool CompareText(XLCellValue value, string actual)
    {
        if (!value.IsText)
            return _comparison == Comparison.NotEqual;

        return _comparison switch
        {
            Comparison.Equal => Wildcard.Matches(actual.AsSpan(), value.GetText().AsSpan()),
            Comparison.NotEqual => !Wildcard.Matches(actual.AsSpan(), value.GetText().AsSpan()),
            Comparison.LessThan => _culture.CompareInfo.Compare(value.GetText(), actual) < 0,
            Comparison.LessOrEqualTo => _culture.CompareInfo.Compare(value.GetText(), actual) <= 0,
            Comparison.GreaterThan => _culture.CompareInfo.Compare(value.GetText(), actual) > 0,
            Comparison.GreaterOrEqualTo => _culture.CompareInfo.Compare(value.GetText(), actual) >= 0,
            _ => throw new UnreachableException()
        };
    }

    private bool CompareError(XLCellValue value, XLError actual)
    {
        if (!value.IsError)
            return _comparison == Comparison.NotEqual;

        return Compare(value.GetError().CompareTo(actual));
    }

    private bool Compare(int cmp)
    {
        return _comparison switch
        {
            Comparison.Equal => cmp == 0,
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
        Equal,
        NotEqual,
        LessThan,
        LessOrEqualTo,
        GreaterThan,
        GreaterOrEqualTo,
    }
}

