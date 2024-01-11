using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class for holding <see cref="XLRichText"/> in a <see cref="SharedStringTable"/>.
    /// It's immutable (keys in reverse dictionary can't change) and more memory efficient
    /// than mutable rich text.
    /// </summary>
    [DebuggerDisplay("{Text}")]
    internal sealed class XLImmutableRichText : IEquatable<XLImmutableRichText>
    {
        private readonly RichTextRun[] _runs;
        private readonly PhoneticRun[] _phoneticRuns;
        
        private XLImmutableRichText(string text, RichTextRun[] runs, PhoneticRun[] phoneticRuns, PhoneticProperties? phoneticsProps)
        {
            Text = text;
            _runs = runs;
            _phoneticRuns = phoneticRuns;
            PhoneticsProperties = phoneticsProps;
        }

        /// <summary>
        /// A text of a whole rich text, without styling.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Individual rich text runs that make up the <see cref="Text"/>, in ascending order, non-overlapping.
        /// </summary>
        public IReadOnlyList<RichTextRun> Runs => _runs;

        /// <summary>
        /// All phonetics runs of rich text. Empty array, if no phonetic run. In ascending order, non-overlapping.
        /// </summary>
        public IReadOnlyList<PhoneticRun> PhoneticRuns => _phoneticRuns;

        /// <summary>
        /// Properties used to display phonetic runs.
        /// </summary>
        public PhoneticProperties? PhoneticsProperties { get; }

        public bool Equals(XLImmutableRichText? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return Text == other.Text &&
                   _runs.SequenceEqual(other._runs) &&
                   _phoneticRuns.SequenceEqual(other._phoneticRuns) &&
                   Nullable.Equals(PhoneticsProperties, other.PhoneticsProperties);
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as XLImmutableRichText);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Text.GetHashCode();
                hashCode = (hashCode * 397) ^ PhoneticsProperties.GetHashCode();
                foreach (var phoneticRun in _phoneticRuns)
                    hashCode = (hashCode * 397) ^ phoneticRun.GetHashCode();

                foreach (var run in _runs)
                    hashCode = (hashCode * 397) ^ run.GetHashCode();

                return hashCode;
            }
        }

        internal string GetRunText(RichTextRun run) => Text.Substring(run.StartIndex, run.Length);

        /// <summary>
        /// Create an immutable rich text with same content as the original <paramref name="formattedText"/>.
        /// </summary>
        internal static XLImmutableRichText Create<T>(XLFormattedText<T> formattedText)
        {
            var text = formattedText.Text;
            var runs = new RichTextRun[formattedText.Count];
            var runIdx = 0;
            var charStartIdx = 0;
            foreach (var richString in formattedText)
            {
                runs[runIdx++] = new RichTextRun(richString, charStartIdx, richString.Text.Length);
                charStartIdx += richString.Text.Length;
            }

            PhoneticRun[] phoneticRuns;
            PhoneticProperties? phoneticProps;
            if (formattedText.HasPhonetics)
            {
                var rtPhonetics = formattedText.Phonetics;
                phoneticRuns = new PhoneticRun[rtPhonetics.Count];
                var phoneticRunIdx = 0;
                var prevPhoneticEndIdx = 0;
                foreach (var phonetic in formattedText.Phonetics)
                {
                    if (phonetic.Start >= text.Length)
                        throw new ArgumentException("Phonetic run start index must be within the text boundaries.");

                    if (phonetic.End > text.Length)
                        throw new ArgumentException("Phonetic run end index must be at most length of a text.");

                    if (phonetic.Start < prevPhoneticEndIdx)
                        throw new ArgumentException("Phonetic runs must be in ascending order and can't overlap.");

                    phoneticRuns[phoneticRunIdx++] = new PhoneticRun(phonetic.Text, phonetic.Start, phonetic.End);
                    prevPhoneticEndIdx = phonetic.End;
                }

                phoneticProps = new PhoneticProperties(formattedText.Phonetics);
            }
            else
            {
                phoneticRuns = Array.Empty<PhoneticRun>();
                phoneticProps = null;
            }

            return new XLImmutableRichText(text, runs, phoneticRuns, phoneticProps);
        }
        
        internal readonly struct RichTextRun : IEquatable<RichTextRun>
        {
            internal readonly int StartIndex;
            internal readonly int Length;
            internal readonly XLFontValue Font;

            internal RichTextRun(XLRichString richString, int startIndex, int length)
            {
                var key = XLFont.GenerateKey(richString);
                Font = XLFontValue.FromKey(ref key);
                StartIndex = startIndex;
                Length = length;
            }

            public bool Equals(RichTextRun other)
            {
                return StartIndex == other.StartIndex && Length == other.Length && Font.Equals(other.Font);
            }

            public override bool Equals(object? obj)
            {
                return obj is RichTextRun other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hashCode = StartIndex;
                    hashCode = (hashCode * 397) ^ Length;
                    hashCode = (hashCode * 397) ^ Font.GetHashCode();
                    return hashCode;
                }
            }
        }

        /// <summary>
        /// Phonetic runs can't overlap and must be in order (i.e. start index must be ascending).
        /// </summary>
        internal readonly struct PhoneticRun
        {
            /// <summary>
            /// Text that is displayed above a segment indicating how should segment be read.
            /// </summary>
            internal readonly string Text;

            /// <summary>
            /// Starting index of displayed phonetic (first character is 0).
            /// </summary>
            internal readonly int StartIndex;

            /// <summary>
            /// End index, excluding (the last index is a length of the rich text).
            /// </summary>
            internal readonly int EndIndex;

            public PhoneticRun(string text, int startIndex, int endIndex)
            {
                if (text.Length == 0)
                    throw new ArgumentException("Phonetic run text can't be empty.", nameof(text));

                if (startIndex < 0)
                    throw new ArgumentException("Start index index must be greater than 0.", nameof(startIndex));

                if (startIndex >= endIndex)
                    throw new ArgumentException("Start index must be less than end index.", nameof(endIndex));

                Text = text;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            public bool Equals(PhoneticRun other)
            {
                return Text == other.Text && StartIndex == other.StartIndex && EndIndex == other.EndIndex;
            }

            public override bool Equals(object? obj)
            {
                return obj is PhoneticRun other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hashCode = Text.GetHashCode();
                    hashCode = (hashCode * 397) ^ StartIndex;
                    hashCode = (hashCode * 397) ^ EndIndex;
                    return hashCode;
                }
            }
        }

        /// <summary>
        /// Properties of phonetic runs. All phonetic runs of a rich text have same font and other properties.
        /// </summary>
        internal readonly struct PhoneticProperties
        {
            /// <summary>
            /// Font used for text of phonetic runs. All phonetic runs use same font. There can be no phonetic runs,
            /// but with specified font (e.g. the mutable API has only specified font, but no text yet).
            /// </summary>
            public readonly XLFontValue Font;

            /// <summary>
            /// Type of phonetics. Default is <see cref="XLPhoneticType.FullWidthKatakana"/>
            /// </summary>
            public readonly XLPhoneticType Type;

            /// <summary>
            /// Alignment of phonetics. Default is <see cref="XLPhoneticAlignment.Left"/>
            /// </summary>
            public readonly XLPhoneticAlignment Alignment;

            public PhoneticProperties(XLPhonetics rtPhonetics)
            {
                var fontKey = XLFont.GenerateKey(rtPhonetics);
                Font = XLFontValue.FromKey(ref fontKey);
                Type = rtPhonetics.Type;
                Alignment = rtPhonetics.Alignment;
            }

            public bool Equals(PhoneticProperties other)
            {
                return Font.Equals(other.Font) && Type == other.Type && Alignment == other.Alignment;
            }

            public override bool Equals(object? obj)
            {
                return obj is PhoneticProperties other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hashCode = Font.GetHashCode();
                    hashCode = (hashCode * 397) ^ (int)Type;
                    hashCode = (hashCode * 397) ^ (int)Alignment;
                    return hashCode;
                }
            }
        }
    }
}
