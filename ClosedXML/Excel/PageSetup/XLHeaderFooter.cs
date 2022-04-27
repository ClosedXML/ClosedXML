using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLHeaderFooter: IXLHeaderFooter
    {

        public XLHeaderFooter(XLWorksheet worksheet)
        {
            Worksheet = worksheet;
            Left = new XLHFItem(this);
            Right = new XLHFItem(this);
            Center = new XLHFItem(this);
            SetAsInitial();
        }

        public XLHeaderFooter(XLHeaderFooter defaultHF, XLWorksheet worksheet)
        {
            Worksheet = worksheet;
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

        public string GetText(XLHFOccurrence occurrence)
        {
            //if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

            var retVal = string.Empty;
            var leftText = Left.GetText(occurrence);
            var centerText = Center.GetText(occurrence);
            var rightText = Right.GetText(occurrence);
            retVal += leftText.Length > 0 ? "&L" + leftText : string.Empty;
            retVal += centerText.Length > 0 ? "&C" + centerText : string.Empty;
            retVal += rightText.Length > 0 ? "&R" + rightText : string.Empty;
            if (retVal.Length > 255)
            {
                throw new ArgumentOutOfRangeException("Headers and Footers cannot be longer than 255 characters (including style markups)");
            }

            return retVal;
        }

        private Dictionary<XLHFOccurrence, string> innerTexts = new Dictionary<XLHFOccurrence, string>();
        internal void SetInnerText(XLHFOccurrence occurrence, string text)
        {
            var parsedElements = ParseFormattedHeaderFooterText(text);

            if (parsedElements.Any(e => e.Position == 'L'))
            {
                Left.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'L').Select(e => e.Text).ToArray()), occurrence);
            }

            if (parsedElements.Any(e => e.Position == 'C'))
            {
                Center.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'C').Select(e => e.Text).ToArray()), occurrence);
            }

            if (parsedElements.Any(e => e.Position == 'R'))
            {
                Right.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'R').Select(e => e.Text).ToArray()), occurrence);
            }

            innerTexts[occurrence] = text;
        }

        private struct ParsedHeaderFooterElement
        {
            public char Position;
            public string Text;
        }

        private static IEnumerable<ParsedHeaderFooterElement> ParseFormattedHeaderFooterText(string text)
        {
            bool IsAtPositionIndicator(int i) => i < text.Length - 1 && text[i] == '&' && new char[] { 'L', 'C', 'R' }.Contains(text[i + 1]);

            var parsedElements = new List<ParsedHeaderFooterElement>();
            var currentPosition = 'L'; // default is LEFT
            var hfElement = "";

            for (var i = 0; i < text.Length; i++)
            {
                if (IsAtPositionIndicator(i))
                {
                    if (!string.IsNullOrEmpty(hfElement))
                    {
                        parsedElements.Add(new ParsedHeaderFooterElement()
                    {
                        Position = currentPosition,
                        Text = hfElement
                    });
                    }

                    currentPosition = text[i + 1];
                    i += 2;
                    hfElement = "";
                }

                if (i < text.Length)
                {
                    if (IsAtPositionIndicator(i))
                    {
                        i--;
                    }
                    else
                    {
                        hfElement += text[i];
                    }
                }
            }

            if (!string.IsNullOrEmpty(hfElement))
            {
                parsedElements.Add(new ParsedHeaderFooterElement()
                {
                    Position = currentPosition,
                    Text = hfElement
                });
            }

            return parsedElements;
        }

        private Dictionary<XLHFOccurrence, string> _initialTexts;

        private bool _changed;
        internal bool Changed
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
