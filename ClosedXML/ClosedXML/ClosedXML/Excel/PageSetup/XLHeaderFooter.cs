using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLHeaderFooter: IXLHeaderFooter
    {
        public XLHeaderFooter()
        {
            Left = new XLHFItem();
            Right = new XLHFItem();
            Center = new XLHFItem();
        }

        public XLHeaderFooter(XLHeaderFooter defaultHF)
        {
            defaultHF.innerTexts.ForEach(kp => innerTexts.Add(kp.Key, kp.Value));
            Left = new XLHFItem(defaultHF.Left as XLHFItem);
            Center = new XLHFItem(defaultHF.Center as XLHFItem);
            Right = new XLHFItem(defaultHF.Right as XLHFItem);
        }

        public IXLHFItem Left { get; private set; }
        public IXLHFItem Center { get; private set; }
        public IXLHFItem Right { get; private set; }

        public String GetText(XLHFOccurrence occurrence)
        {
            if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

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
        internal String SetInnerText(XLHFOccurrence occurrence, String text)
        { 
            if (innerTexts.ContainsKey(occurrence))
                innerTexts[occurrence] = text;
            else
                innerTexts.Add(occurrence, text);

            return innerTexts[occurrence];
        }
    }
}
