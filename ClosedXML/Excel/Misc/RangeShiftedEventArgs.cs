using System;

namespace ClosedXML.Excel.Misc
{
    internal class RangeShiftedEventArgs : EventArgs
    {
        public RangeShiftedEventArgs(XLRange range, int shifted)
        {
            this.Range = range;
            this.Shifted = shifted;
        }

        public new static RangeShiftedEventArgs Empty
        {
            get { return new RangeShiftedEventArgs(null, 0); }
        }

        public XLRange Range { get; private set; }
        public int Shifted { get; private set; }
    }
}
