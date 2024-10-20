#nullable disable

using System;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A wildcard is at most 255 chars long text. It can contain <c>*</c> which indicates any number characters (including zero)
    /// and <c>?</c> which indicates any single character. If you need to find <c>*</c> or <c>?</c> in a text, prefix them with
    /// an escape character <c>~</c>.
    /// </summary>
    internal readonly struct Wildcard
    {
        private readonly String _pattern;

        public Wildcard(String pattern)
        {
            _pattern = pattern;
        }

        /// <summary>
        /// Search for the wildcard anywhere in the text.
        /// </summary>
        /// <param name="text">Text used to search for a pattern.</param>
        /// <returns>zero-based index of a first character in a text that matches to a pattern or -1, if match wasn't found.</returns>
        public int Search(ReadOnlySpan<char> text)
        {
            return Search(_pattern.AsSpan(), text).StartIndex;
        }

        /// <summary>
        /// Does pattern matches whole text?
        /// </summary>
        public static bool Matches(ReadOnlySpan<char> pattern, ReadOnlySpan<char> text)
        {
            var (startIndex, endIndex) = Search(pattern, text);
            if (startIndex != 0)
                return false;

            if (endIndex != text.Length)
                return false;

            return true;
        }

        private static (int StartIndex, int EndIndex) Search(ReadOnlySpan<char> pattern, ReadOnlySpan<char> text)
        {
            // Excel limits pattern size to 255, likely to avoid performance problems due to backtracking.
            if (pattern.Length > 255)
                return (-1, 0);

            // Check to remove trailing escape ~
            if (pattern.Length >= 2 && pattern[pattern.Length - 1] == '~' && pattern[pattern.Length - 2] != '~')
                pattern = pattern.Slice(0, pattern.Length - 1);

            // pattern index should be an index of first character of a segment. Segments start and ends with a star.
            var patternIdx = 0;

            // Index of a first segment that was recognized in the text. -1 indicates not found.
            var firstSegmentStartIdx = -1;

            // Skip the first stars in the pattern. The Search method searches within the text and it doesn't have
            // to start at the beginning, thus stars are irrelevant for Search method.
            while (patternIdx < pattern.Length && pattern[patternIdx] == '*')
            {
                patternIdx++;

                // If the pattern starts with a star, it must start at the beginning of a text
                firstSegmentStartIdx = 0;
            }

            if (patternIdx >= pattern.Length)
            {
                // Whole pattern consists of * wildcards, any text satisfies the pattern.
                return (0, text.Length);
            }

            // Text index points to the first char of yet unprocessed text. As pattern segments are matched in a text, text index increases,
            // so the same substring of a text can't be used to match two segments.
            var textIdx = 0;

            // Because of escapes, we can't just check end character for star, we have to track it. Updated for every segment is processed.
            var endsWithStar = false;

            while (patternIdx < pattern.Length)
            {
                if (textIdx >= text.Length)
                {
                    // There is still a non-star pattern, but text has ended - there is no way we can find the last segment in the text.
                    return (-1, 0);
                }

                endsWithStar = false;

                // Each loop is searching only for a specific a segment in a text that hasn't yet been processed.
                // Segment pattern starts after previous star wildcard and ends before next star wildcard/end of pattern.
                var segmentEnd = patternIdx;
                for (; segmentEnd < pattern.Length; ++segmentEnd)
                {
                    var segmentChar = pattern[segmentEnd];
                    if (segmentChar == '*')
                    {
                        endsWithStar = true;
                        break;
                    }

                    if (segmentChar == '~' && segmentEnd + 1 < pattern.Length)
                        segmentEnd++;
                }

                var segmentLength = segmentEnd - patternIdx;
                var segmentPattern = pattern.Slice(patternIdx, segmentLength);

                // Search only in a text that hasn't yet been used to match some previous segment pattern.
                var textAfterPrevSegment = text.Slice(textIdx);
                var patternPosInSegment = SearchSegment(segmentPattern, textAfterPrevSegment);
                if (patternPosInSegment.TextStartIdx < 0)
                {
                    // Segment pattern is not present in the text -> whole pattern isn't in the text
                    return (-1, 0);
                }

                if (firstSegmentStartIdx < 0)
                    firstSegmentStartIdx = patternPosInSegment.TextStartIdx;

                patternIdx += segmentLength;
                textIdx += patternPosInSegment.TextAfterLastIdx;

                // Skip stars between segments. Due to backtracking, they are rather irrelevant.
                while (patternIdx < pattern.Length && pattern[patternIdx] == '*')
                {
                    endsWithStar = true;
                    patternIdx++;
                }
            }

            if (endsWithStar)
                textIdx = text.Length;

            return (firstSegmentStartIdx, textIdx);
        }
        
        /// <summary>
        /// Check if a segment can be found in a text.
        /// </summary>
        /// <param name="segmentPattern">Non-empty pattern without <c>*</c> wildcard, though it can contain ? and escaped *.</param>
        /// <param name="text">Non-empty text.</param>
        /// <returns>First index of the pattern in the <paramref name="text"/> or -1 if not found.</returns>
        private static (int TextStartIdx, int TextAfterLastIdx) SearchSegment(ReadOnlySpan<char> segmentPattern, ReadOnlySpan<char> text)
        {
            var textBacktrackStart = 0;
            var patternIdx = 0;
            var textIdx = 0;

            while (patternIdx < segmentPattern.Length)
            {
                if (textIdx >= text.Length)
                {
                    // There is unmatched star-less pattern, but text is already over -> impossible to match the rest of a segment pattern
                    return (-1, 0);
                }

                var patternChar = segmentPattern[patternIdx++];
                if (patternChar == '?')
                {
                    textIdx++;
                    continue;
                }

                if (patternChar == '~')
                {
                    if (patternIdx < segmentPattern.Length)
                    {
                        patternChar = segmentPattern[patternIdx++];
                    }
                    else
                    {
                        // There is only escape char, but the pattern has ended. Thus we matched.
                        return (textBacktrackStart, textIdx);
                    }
                }

                var textChar = text[textIdx++];
                var sameChar = textChar == patternChar ||
                               char.ToUpperInvariant(textChar) == char.ToUpperInvariant(patternChar);
                if (!sameChar)
                {
                    // Text and pattern don't match - we need to backtrack.
                    if (++textBacktrackStart >= text.Length)
                    {
                        return (-1, 0);
                    }

                    textIdx = textBacktrackStart;
                    patternIdx = 0;
                }
            }

            return (textBacktrackStart, textIdx);
        }
    }
}
