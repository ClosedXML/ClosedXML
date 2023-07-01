namespace ClosedXML.Excel
{
    internal struct XLMiscSliceContent
    {
        // Must be as flag for inline string, so the default value is false => ShareString is true by default 
        private bool _inlineString;

        internal bool ShareString
        {
            get => !_inlineString;
            set => _inlineString = !value;
        }

        internal XLComment? Comment { get; set; }

        internal XLHyperlink? Hyperlink { get; set; }

        internal uint? CellMetaIndex { get; set; }

        internal uint? ValueMetaIndex { get; set; }

        internal bool SettingHyperlink { get; set; }

        internal bool HasPhonetic { get; set; }
    }
}
