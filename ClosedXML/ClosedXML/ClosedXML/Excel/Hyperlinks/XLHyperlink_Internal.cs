using System;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        internal XLHyperlink()
        { 

        }

        internal XLHyperlink(XLHyperlink hyperlink)
        {
            externalAddress = hyperlink.externalAddress;
            internalAddress = hyperlink.internalAddress;
            Tooltip = hyperlink.Tooltip;
            IsExternal = hyperlink.IsExternal;
        }

        internal void SetValues(String address, String tooltip)
        {
            Tooltip = tooltip;
            if (address[0] == '.')
            {
                externalAddress = new Uri(address, UriKind.Relative);
                IsExternal = true;
            }
            else
            {
                Uri uri;
                if(Uri.TryCreate(address, UriKind.Absolute, out uri))
                {
                    externalAddress = uri;
                    IsExternal = true;
                }
                else
                {
                    internalAddress = address;
                    IsExternal = false;    
                }
            }
        }

        internal void SetValues(Uri uri, String tooltip)
        {
            Tooltip = tooltip;
            externalAddress = uri;
            IsExternal = true;
        }

        internal void SetValues(IXLCell cell, String tooltip)
        {
            Tooltip = tooltip;
            internalAddress = cell.Address.ToString();
            IsExternal = false;
        }

        internal void SetValues(IXLRangeBase range, String tooltip)
        {
            Tooltip = tooltip;
            internalAddress = range.RangeAddress.ToString();
            IsExternal = false;
        }

        internal XLWorksheet Worksheet { get; set; }
    }
}
