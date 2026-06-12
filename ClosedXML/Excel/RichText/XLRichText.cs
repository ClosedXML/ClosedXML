using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An API object for manipulating rich text. Every time it is changed, it calls
    /// <see cref="OnContentChanged"/> to project changes back to the <see cref="SharedStringTable"/>.
    /// </summary>
    internal class XLRichText : XLFormattedText<IXLRichText>, IXLRichText
    {
        // Should be set as the last thing in ctor to prevent firing changes to immutable rich text during ctor
        private readonly XLCell? _cell;

        /// <summary>
        /// Copy ctor to return user modifiable rich text from an immutable rich text stored
        /// in the shared string table.
        /// </summary>
        internal XLRichText(XLCell cell, XLFontFormatValue defaultFont, XLImmutableRichText original)
            : base(defaultFont, cell.Worksheet.Workbook.Styles)
        {
            foreach (var originalRun in original.Runs)
            {
                var runText = original.GetRunText(originalRun);
                AddText(new XLRichString(runText, originalRun.Font, this, Styles, OnContentChanged));
            }

            var hasPhonetics = original.PhoneticRuns.Count > 0 || original.PhoneticsProperties.HasValue;
            if (hasPhonetics)
            {
                XLPhonetics phonetics;
                if (original.PhoneticsProperties.HasValue)
                {
                    var originalProps = original.PhoneticsProperties.Value;
                    phonetics = new XLPhonetics(originalProps.Font, defaultFont, Styles, OnContentChanged)
                    {
                        Type = originalProps.Type,
                        Alignment = originalProps.Alignment
                    };
                }
                else
                {
                    phonetics = new XLPhonetics(defaultFont, defaultFont, Styles, OnContentChanged);
                }

                foreach (var phoneticRun in original.PhoneticRuns)
                    phonetics.Add(phoneticRun.Text, phoneticRun.StartIndex, phoneticRun.EndIndex);

                Phonetics = phonetics;
            }

            // TODO Styles: Convert to a factory method. The cell is set at the end to avoid false change trigger. Refactor so it's not needed anymore
            Container = this;
            _cell = cell;
        }

        internal XLRichText(XLCell cell, XLFontFormatValue defaultFont, String text)
            : this(cell, defaultFont)
        {
            AddText(new XLRichString(text, defaultFont, this, Styles, OnContentChanged));
        }

        internal XLRichText(XLCell cell, XLFontFormatValue defaultFont)
            : base(defaultFont, cell.Worksheet.Workbook.Styles)
        {
            Container = this;
            _cell = cell;
        }

        protected override void OnContentChanged()
        {
            // The rich text is still being created
            if (_cell is null)
                return;

            if (_cell.DataType != XLDataType.Text || !_cell.HasRichText)
                throw new InvalidOperationException("The rich text isn't a content of a cell.");

            _cell.SetOnlyValue(Text);
            var point = _cell.Point;
            var richText = XLImmutableRichText.Create(this);
            _cell.Worksheet.Internals.CellsCollection.ValueSlice.SetRichText(point, richText);
        }
    }
}
