using System;

namespace ClosedXML.Excel
{
    internal class XLPhonetic: IXLPhonetic
    {
        public XLPhonetic(String text, Int32 start, Int32 end)
        {
            Text = text;
            Start = start;
            End = end;
        }
        public String Text { get; set; }
        public Int32 Start { get; set; }
        public Int32 End { get; set; }

        public bool Equals(IXLPhonetic other)
        {
            if (other == null)
                return false;

            return Text == other.Text && Start == other.Start && End == other.End;
        }
    }
}
