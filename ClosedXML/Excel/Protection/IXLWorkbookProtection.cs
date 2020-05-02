// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    public interface IXLWorkbookProtection : ICloneable
    {
        Algorithm Algorithm { get; }

        XLWorkbookProtectionElements AllowedElements { get; set; }

        Boolean IsPasswordProtected { get; }

        Boolean IsProtected { get; }

        /// <summary>
        /// Adds the workbook protection element to the list of allowed elements.
        /// Beware that if you pass through <see cref="XLWorkbookProtectionElements.None" />, this will have no effect.
        /// </summary>
        /// <param name="element">The workbook protection element to add</param>
        /// <param name="allowed">Set to <c>true</c> to allow the element or <c>false</c> to disallow the element</param>
        /// <returns>The current workbook protection</returns>
        IXLWorkbookProtection AllowElement(XLWorkbookProtectionElements element, Boolean allowed = true);

        IXLWorkbookProtection AllowEverything();

        IXLWorkbookProtection AllowNone();

        IXLWorkbookProtection CopyFrom(IXLWorkbookProtection workbookProtection);

        /// <summary>
        /// Removes the workbook protection element to the list of allowed elements.
        /// Beware that if you pass through <see cref="XLWorkbookProtectionElements.None" />, this will have no effect.
        /// </summary>
        /// <param name="element">The workbook protection element to remove</param>
        /// <returns>The current workbook protection</returns>
        IXLWorkbookProtection DisallowElement(XLWorkbookProtectionElements element);

        IXLWorkbookProtection Protect();

        IXLWorkbookProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows);

        IXLWorkbookProtection Unprotect();

        IXLWorkbookProtection Unprotect(String password);
    }
}
