﻿namespace ClosedXML.Excel
{
    internal struct XLMiscSliceContent
    {
        internal XLComment? Comment { get; set; }

        internal XLHyperlink? Hyperlink { get; set; }

        internal uint? CellMetaIndex { get; set; }

        internal uint? ValueMetaIndex { get; set; }

        internal bool SettingHyperlink { get; set; }

        internal bool HasPhonetic { get; set; }
    }
}
