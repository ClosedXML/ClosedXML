using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        public Boolean IsExternal { get; set; }

        public XLHyperlink(String address)
        {
            SetValues(address, String.Empty);
        }
        public XLHyperlink(String address, String tooltip)
        {
            SetValues(address, tooltip);
        }

        public XLHyperlink(IXLCell cell)
        {
            SetValues(cell, String.Empty);
        }
        public XLHyperlink(IXLCell cell, String tooltip)
        {
            SetValues(cell, tooltip);
        }

        public XLHyperlink(IXLRangeBase range)
        {
            SetValues(range, String.Empty);
        }
        public XLHyperlink(IXLRangeBase range, String tooltip)
        {
            SetValues(range, tooltip);
        }

        public XLHyperlink(Uri uri)
        {
            SetValues(uri, String.Empty);
        }
        public XLHyperlink(Uri uri, String tooltip)
        {
            SetValues(uri, tooltip);
        }

        private Uri externalAddress;
        public Uri ExternalAddress 
        {
            get
            {
                if (IsExternal)
                    return externalAddress;
                else
                    return null;
            }
            set
            {
                externalAddress = value;
                IsExternal = true;
            }
        }

        public IXLCell Cell { get; internal set; }

        private String internalAddress;
        public String InternalAddress
        {
            get
            {
                if (IsExternal)
                {
                    return null;
                }
                else
                {
                    if (internalAddress.Contains('!'))
                    {
                        if (internalAddress[0] != '\'')
                            return String.Format("'{0}'!{1}", internalAddress.Substring(0, internalAddress.IndexOf('!')), internalAddress.Substring(internalAddress.IndexOf('!') + 1));
                        else
                            return internalAddress;
                    }
                    else
                    {
                        return String.Format("'{0}'!{1}", Worksheet.Name, internalAddress);
                    }
                }
            }
            set
            {
                internalAddress = value;
                IsExternal = false;
            }
        }

        public String Tooltip { get; set; }
        public void Delete()
        {
            if (Cell != null)
            {
                Worksheet.Hyperlinks.Delete(Cell.Address);
                if (Cell.Style.Font.FontColor.Equals(XLColor.FromTheme(XLThemeColor.Hyperlink)))
                    Cell.Style.Font.FontColor = Worksheet.Style.Font.FontColor;

                if (Cell.Style.Font.Underline != Worksheet.Style.Font.Underline)
                    Cell.Style.Font.Underline = Worksheet.Style.Font.Underline;
            }
        }
    }
}
