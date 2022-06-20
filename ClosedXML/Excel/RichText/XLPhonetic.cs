namespace ClosedXML.Excel
{
    internal class XLPhonetic: IXLPhonetic
    {
        public XLPhonetic(string text, int start, int end)
        {
            Text = text;
            Start = start;
            End = end;
        }
        public string Text { get; set; }
        public int Start { get; set; }
        public int End { get; set; }

        public bool Equals(IXLPhonetic other)
        {
            if (other == null)
            {
                return false;
            }

            return Text == other.Text && Start == other.Start && End == other.End;
        }
    }
}
