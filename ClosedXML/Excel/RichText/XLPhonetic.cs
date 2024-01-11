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
        public String Text { get; }
        public Int32 Start { get; }
        public Int32 End { get; }

        public bool Equals(IXLPhonetic? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return Text == other.Text && Start == other.Start && End == other.End;
        }
    }
}
