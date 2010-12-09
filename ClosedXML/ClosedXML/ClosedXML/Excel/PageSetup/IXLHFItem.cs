using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLHFPredefinedText
    { 
        PageNumber, NumberOfPages, Date, Time, FullPath, Path, File, SheetName
    }
    public enum XLHFOccurrence
    { 
        AllPages, OddPages, EvenPages, FirstPage
    }

    public interface IXLHFItem
    {
        String GetText(XLHFOccurrence occurrence);
        void AddText(String text);
        void AddText(XLHFPredefinedText predefinedText);
        void AddText(String text, XLHFOccurrence occurrence);
        void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence);
        void AddText(String text, XLHFOccurrence occurrence, IXLFont xlFont);
        void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence, IXLFont xlFont);
        void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);
    }
}
