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
        void AddText(String text, XLHFOccurrence occurrence = XLHFOccurrence.AllPages, IXLFont xlFont = null);
        void AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence = XLHFOccurrence.AllPages, IXLFont xlFont = null);
        void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);
    }
}
