using System;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// A font name, two font names are equal when they are case insensitive equal. It is a custom
/// class because that way <see cref="XLFontFormat"/> and other structures don't have to implement
/// custom hash code and equality methods.
/// </summary>
internal readonly record struct XLFontName
{
    private const StringComparison Comparison = StringComparison.OrdinalIgnoreCase;

    private XLFontName(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException(nameof(text));

        Text = text;
    }

    public string Text { get; }

    public override int GetHashCode()
    {
        return Text.GetHashCode(Comparison);
    }

    public bool Equal(XLFontName other)
    {
        return string.Equals(Text, other.Text, Comparison);
    }

    [return: NotNullIfNotNull(nameof(text))]
    public static implicit operator XLFontName?(string? text) => !string.IsNullOrWhiteSpace(text) ? new XLFontName(text) : null;
}
