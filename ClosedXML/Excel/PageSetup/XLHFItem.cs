﻿using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLHFItem : IXLHFItem
    {
        private readonly XLWorksheet _worksheet;
        public XLHFItem(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }
        public XLHFItem(XLHFItem defaultHFItem, XLWorksheet worksheet)
            :this(worksheet)
        {
            defaultHFItem.texts.ForEach(kp => texts.Add(kp.Key, kp.Value));
        }
        private readonly Dictionary<XLHFOccurrence, List<XLHFText>> texts = new Dictionary<XLHFOccurrence, List<XLHFText>>();
        public String GetText(XLHFOccurrence occurrence)
        {
            var sb = new StringBuilder();
            if(texts.ContainsKey(occurrence))
            {
                foreach (var hfText in texts[occurrence])
                    sb.Append(hfText.GetHFText(sb.ToString()));
            }

            return sb.ToString();
        }

        public IXLRichString AddText(String text)
        {
            return AddText(text, XLHFOccurrence.AllPages);
        }
        public IXLRichString AddText(XLHFPredefinedText predefinedText)
        {
            return AddText(predefinedText, XLHFOccurrence.AllPages);
        }

        public IXLRichString AddText(String text, XLHFOccurrence occurrence)
        {
            XLRichString richText = new XLRichString(text, XLWorkbook.DefaultStyle.Font, this);

            var hfText = new XLHFText(richText, _worksheet);
            if (occurrence == XLHFOccurrence.AllPages)
            {
                AddTextToOccurrence(hfText, XLHFOccurrence.EvenPages);
                AddTextToOccurrence(hfText, XLHFOccurrence.FirstPage);
                AddTextToOccurrence(hfText, XLHFOccurrence.OddPages);
            }
            else
            {
                AddTextToOccurrence(hfText, occurrence);
            }

            return richText;
        }

        public IXLRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

        public IXLRichString AddImage(String imagePath, XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
        {
            throw new NotImplementedException();
        }

        private void AddTextToOccurrence(XLHFText hfText, XLHFOccurrence occurrence)
        {
            if (texts.ContainsKey(occurrence))
                texts[occurrence].Add(hfText);
            else
                texts.Add(occurrence, new List<XLHFText> { hfText });
        }

        public IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence)
        {
            String hfText;
            switch (predefinedText)
            {
                case XLHFPredefinedText.PageNumber: hfText = "&P"; break;
                case XLHFPredefinedText.NumberOfPages : hfText = "&N"; break;
                case XLHFPredefinedText.Date : hfText = "&D"; break;
                case XLHFPredefinedText.Time : hfText = "&T"; break;
                case XLHFPredefinedText.Path : hfText = "&Z"; break;
                case XLHFPredefinedText.File : hfText = "&F"; break;
                case XLHFPredefinedText.SheetName : hfText = "&A"; break;
                case XLHFPredefinedText.FullPath: hfText = "&Z&F"; break;
                default: throw new NotImplementedException();
            }
            return AddText(hfText, occurrence);
        }

        public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
        {
            if (occurrence == XLHFOccurrence.AllPages)
            {
                ClearOccurrence(XLHFOccurrence.EvenPages);
                ClearOccurrence(XLHFOccurrence.FirstPage);
                ClearOccurrence(XLHFOccurrence.OddPages);
            }
            else
            {
                ClearOccurrence(occurrence);
            }
        }

        private void ClearOccurrence(XLHFOccurrence occurrence)
        {
            if (texts.ContainsKey(occurrence))
                texts.Remove(occurrence);
        }
    }
}
