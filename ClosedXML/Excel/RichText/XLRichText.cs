using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRichText : XLFormattedText<IXLRichText>, IXLRichText
    {
        // Should be set as the last thing in ctor to prevent firing changes to immutable rich text during ctor
        private readonly XLCell? _cell;

        /// <summary>
        /// Copy ctor to return user modifiable rich text from immutable rich text stored
        /// in the shared string table.
        /// </summary>
        public XLRichText(XLCell cell, XLImmutableRichText original)
            : base(cell.Style.Font)
        {
            foreach (var originalRun in original.Runs)
            {
                var runText = original.GetRunText(originalRun);
                AddText(new XLRichString(runText, new XLFont(originalRun.Font.Key), this, OnContentChanged));
            }

            var hasPhonetics = original.PhoneticRuns.Any() || original.PhoneticsProperties.HasValue;
            if (hasPhonetics)
            {
                // Access to phonetics instantiate a new instance.
                var phonetics = Phonetics;
                if (original.PhoneticsProperties.HasValue)
                {
                    var phoneticProps = original.PhoneticsProperties.Value;
                    phonetics.CopyFont(new XLFont(phoneticProps.Font.Key));
                    phonetics.Type = phoneticProps.Type;
                    phonetics.Alignment = phoneticProps.Alignment;
                }

                foreach (var phoneticRun in original.PhoneticRuns)
                    phonetics.Add(phoneticRun.Text, phoneticRun.StartIndex, phoneticRun.EndIndex);
            }

            Container = this;
            _cell = cell;
        }

        public XLRichText(XLCell cell, IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Container = this;
            _cell = cell;
        }

        public XLRichText(XLCell cell, String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
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
            var point = _cell.SheetPoint;
            var richText = XLImmutableRichText.Create(this);
            _cell.Worksheet.Internals.CellsCollection.ValueSlice.SetRichText(point, richText);
        }
    }
}
