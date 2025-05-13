using System;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// A font name, two font names are equal when they are case insensitive equal. It is a custom
/// class because that way <see cref="XLFontFormat"/> and other structures don't have to implement
/// custom hash code and equality methods.
/// </summary>
internal readonly record struct XLFontName : IEquatable<string>
{
    private const StringComparison Comparison = StringComparison.OrdinalIgnoreCase;

    private XLFontName(string text)
    {
        // Spec says at most 31 chars, Excel also tries to repair workbook when value is longer.
        if (string.IsNullOrWhiteSpace(text) || text.Length > 31)
            throw new ArgumentException(nameof(text));

        Text = text;
    }

    public string Text { get; }

    public bool Equals(string other)
    {
        return string.Equals(Text, other, Comparison);
    }

    public override int GetHashCode()
    {
        return Text.GetHashCode(Comparison);
    }

    public bool Equal(XLFontName other)
    {
        return Equals(other.Text);
    }

    [return: NotNullIfNotNull(nameof(text))]
    public static implicit operator XLFontName?(string? text) => !string.IsNullOrWhiteSpace(text) ? new XLFontName(text) : null;
}
