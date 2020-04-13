using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    using System.Linq;

    internal class XLHeaderFooter: IXLHeaderFooter
    {

        public XLHeaderFooter(XLWorksheet worksheet)
        {
            this.Worksheet = worksheet;
            Left = new XLHFItem(this);
            Right = new XLHFItem(this);
            Center = new XLHFItem(this);
            SetAsInitial();
        }

        public XLHeaderFooter(XLHeaderFooter defaultHF, XLWorksheet worksheet)
        {
            this.Worksheet = worksheet;
            defaultHF.innerTexts.ForEach(kp => innerTexts.Add(kp.Key, kp.Value));
            Left = new XLHFItem(defaultHF.Left as XLHFItem, this);
            Center = new XLHFItem(defaultHF.Center as XLHFItem, this);
            Right = new XLHFItem(defaultHF.Right as XLHFItem, this);
            SetAsInitial();
        }

        internal readonly IXLWorksheet Worksheet;

        public IXLHFItem Left { get; private set; }
        public IXLHFItem Center { get; private set; }
        public IXLHFItem Right { get; private set; }

        public String GetText(XLHFOccurrence occurrence)
        {
            //if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

            var retVal = String.Empty;
            var leftText = Left.GetText(occurrence);
            var centerText = Center.GetText(occurrence);
            var rightText = Right.GetText(occurrence);
            retVal += leftText.Length > 0 ? "&L" + leftText : String.Empty;
            retVal += centerText.Length > 0 ? "&C" + centerText : String.Empty;
            retVal += rightText.Length > 0 ? "&R" + rightText : String.Empty;
            if (retVal.Length > 255)
                throw new ArgumentOutOfRangeException("Headers and Footers cannot be longer than 255 characters (including style markups)");
            return retVal;
        }

        private Dictionary<XLHFOccurrence, String> innerTexts = new Dictionary<XLHFOccurrence, String>();
        internal void SetInnerText(XLHFOccurrence occurrence, String text)
        {
            var parsedElements = ParseFormattedHeaderFooterText(text);

            if (parsedElements.Any(e => e.Position == 'L'))
                this.Left.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'L').Select(e => e.Text).ToArray()), occurrence);

            if (parsedElements.Any(e => e.Position == 'C'))
                this.Center.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'C').Select(e => e.Text).ToArray()), occurrence);

            if (parsedElements.Any(e => e.Position == 'R'))
                this.Right.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'R').Select(e => e.Text).ToArray()), occurrence);

            innerTexts[occurrence] = text;
        }

        private struct ParsedHeaderFooterElement
        {
            public char Position;
            public string Text;
        }

        private static IEnumerable<ParsedHeaderFooterElement> ParseFormattedHeaderFooterText(string text)
        {
            Func<int, bool> IsAtPositionIndicator = i => i < text.Length - 1 && text[i] == '&' && (new char[] { 'L', 'C', 'R' }.Contains(text[i + 1]));

            var parsedElements = new List<ParsedHeaderFooterElement>();
            var currentPosition = 'L'; // default is LEFT
            var hfElement = "";

            for (int i = 0; i < text.Length; i++)
            {
                if (IsAtPositionIndicator(i))
                {
                    if ("" != hfElement) parsedElements.Add(new ParsedHeaderFooterElement()
                    {
                        Position = currentPosition,
                        Text = hfElement
                    });

                    currentPosition = text[i + 1];
                    i += 2;
                    hfElement = "";
                }

                if (i < text.Length)
                {
                    if (IsAtPositionIndicator(i))
                        i--;
                    else
                        hfElement += text[i];
                }
            }

            if ("" != hfElement)
                parsedElements.Add(new ParsedHeaderFooterElement()
                {
                    Position = currentPosition,
                    Text = hfElement
                });
            return parsedElements;
        }

        private Dictionary<XLHFOccurrence, String> _initialTexts;

        private Boolean _changed;
        internal Boolean Changed
        {
            get
            {
                return _changed || _initialTexts.Any(it => GetText(it.Key) != it.Value);
            }
            set { _changed = value; }
        }

        internal void SetAsInitial()
        {
            _initialTexts = new Dictionary<XLHFOccurrence, string>();
            foreach (var o in Enum.GetValues(typeof(XLHFOccurrence)).Cast<XLHFOccurrence>())
            {
                _initialTexts.Add(o, GetText(o));
            }
        }


        public IXLHeaderFooter Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
        {
            Left.Clear(occurrence);
            Right.Clear(occurrence);
            Center.Clear(occurrence);
            return this;
        }
    }
}
