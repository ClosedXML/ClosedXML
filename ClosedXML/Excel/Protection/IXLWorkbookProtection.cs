// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    public interface IXLWorkbookProtection : IXLElementProtection<XLWorkbookProtectionElements>
    {
        IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements);

        IXLWorkbookProtection Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements);

        IXLWorkbookProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows);
    }
}
