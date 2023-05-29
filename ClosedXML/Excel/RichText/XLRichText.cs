#nullable disable

using System;

namespace ClosedXML.Excel
{
    internal class XLRichText : XLFormattedText<IXLRichText>, IXLRichText
    {
        private readonly XLCell _cell;

        public XLRichText(XLCell cell, IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Container = this;
            _cell = cell;
            ContentChanged += OnContentChanged;
        }

        public XLRichText(XLCell cell, XLFormattedText<IXLRichText> defaultRichText, IXLFontBase defaultFont)
            : base(defaultRichText, defaultFont)
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

        private void OnContentChanged(object sender, EventArgs e)
        {
            if (_cell.DataType != XLDataType.Text || !_cell.HasRichText || !ReferenceEquals(_cell.GetRichText(), this))
                throw new InvalidOperationException("The rich text isn't a content of a cell.");

            _cell.SetOnlyValue(Text);
        }
    }
}
