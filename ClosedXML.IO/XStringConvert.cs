using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;

namespace ClosedXML.IO;

/// <summary>
/// Class to deal with encoding and decoding XString. XString decoding doesn't depend on the source
/// encoding. XString decoding decodes from unicode codepoints (regardless if UTF-8 or UTF-16)
/// and the decoder replaces XString patterns with decoded codepoints.
/// </summary>
public class XStringConvert
{
    /// <summary>
    /// Decode an XString to normal string.
    /// </summary>
    /// <remarks>
    /// There is a similar method in dotnet <see cref="XmlConvert.DecodeName"/>. That one however
    /// also accepts uppercase X (<c>_X????_</c>) which shouldn't be decoded. Hexadecimal digits are
    /// case insensitive, the <c>x</c> marker is not (<a href="https://github.com/ClosedXML/ClosedXML/issues/1154">#1154</a>).
    /// Another issue is that <see cref="XmlConvert.DecodeName"/> accepts 8 hex digits (e.g.,
    /// <c>_xD83DDE43_</c>). That is also not valid for XString.
    /// </remarks>
    /// <param name="text">Test that might contain XString encoded characters.</param>
    /// <returns>Decoded string.</returns>
    [return: NotNullIfNotNull(nameof(text))]
    public static string? Decode(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        // This method is on a hotpath and shouldn't allocate unless necessary.
        // Do lazy initialization, there might not be any pattern at all.
        StringBuilder? sb = null;

        // An index of next character after last encountered XString pattern.
        // Initial value is 0 because text from that index on is copied.
        var prevSliceNextCharIndex = 0;
        var textSpan = text.AsSpan();
        for (var i = textSpan.IndexOf('_'); i >= 0 && i < text.Length - 6; i = text.IndexOf('_', i + 1))
        {
            if (IsPattern(textSpan, i))
            {
                sb ??= new StringBuilder(text.Length);

                // Append text from last XString splice
                sb.Append(text, prevSliceNextCharIndex, i - prevSliceNextCharIndex);

                // Get hex digits from from _xABCD_ patterns. Polyfill doesn't have allocation-free
                // API, so just decode the hex number.
                var codepoint = GetHexValue(textSpan.Slice(i + 2, 4));

                sb.Append((char)codepoint);

                // Move from opening '_' to closing '_' because we have effectively read all that
                // The loop will add + 1 and moves to the next char of text.
                i += 6;
                prevSliceNextCharIndex = i + 1;
            }
        }

        // Not even one pattern was actually replaced -> return original text
        if (sb is null)
            return text;

        sb.Append(text, prevSliceNextCharIndex, text.Length - prevSliceNextCharIndex);
        return sb.ToString();

        // Does the XString pattern starts at index i?
        static bool IsPattern(ReadOnlySpan<char> input, int i)
        {
            // Reorder to ensure simplest tests are checked first
            return i + 6 < input.Length &&
                   input[i] == '_' &&
                   input[i + 1] == 'x' &&
                   input[i + 6] == '_' &&
                   IsHex(input[i + 2]) &&
                   IsHex(input[i + 3]) &&
                   IsHex(input[i + 4]) &&
                   IsHex(input[i + 5]);
        }

    }

    internal static bool TryGetHexValue(ReadOnlySpan<char> text, out uint value)
    {
        foreach (var c in text)
        {
            if (!IsHex(c))
            {
                value = 0;
                return false;
            }
        }

        value = GetHexValue(text);
        return true;
    }

    private static uint GetHexValue(ReadOnlySpan<char> text)
    {
        if (text.Length > 8)
            throw new ArgumentException();

        var codepoint = 0u;
        foreach (var c in text)
        {
            var hexDigit = (uint)GetHex(c);
            codepoint = (codepoint * 16) + hexDigit;
        }

        return codepoint;
    }

    private static bool IsHex(char c)
    {
        return c is >= '0' and <= '9' ||
               c is >= 'A' and <= 'F' ||
               c is >= 'a' and <= 'f';
    }

    // We already know that c passed the IsHex method.
    private static int GetHex(char c)
    {
        return c switch
        {
            >= 'A' and <= 'F' => c - 'A' + 10,
            >= 'a' and <= 'f' => c - 'a' + 10,
            >= '0' and <= '9' => c - '0',
            _ => throw new UnreachableException()
        };
    }
}
