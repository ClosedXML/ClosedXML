using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLWorkbookProtectionElements
    {
        None = 0,
        Structure = 1 << 0,

        /// <summary>
        /// The Windows option is available only in Excel 2007, Excel 2010, Excel for Mac 2011, and Excel 2016 for Mac. Select the Windows option if you want to prevent users from moving, resizing, or closing the workbook window, or hide/unhide windows.
        /// </summary>
        Windows = 1 << 1,

        Everything = Structure | Windows
    }
}
