#nullable disable

// Keep this file CodeMaid organised and cleaned
using DocumentFormat.OpenXml;
using System;

namespace ClosedXML.Excel
{
    public class SaveOptions
    {
        public SaveOptions()
        {
#if DEBUG
            this.ValidatePackage = true;
#else
            this.ValidatePackage = false;
#endif
        }

        public Boolean ConsolidateConditionalFormatRanges { get; set; } = true;
        public Boolean ConsolidateDataValidationRanges { get; set; } = true;

        /// <summary>
        /// Evaluate a cells with a formula and save the calculated value along with the formula.
        /// <list type="bullet">
        /// <item>
        ///   True - formulas are evaluated and the calculated values are saved to the file.
        ///   If evaluation of a formula throws an exception, value is not saved but file is still saved.
        /// </item>
        /// <item>
        ///   False (default) - formulas are not evaluated and the formula cells don't have their values saved to the file.
        /// </item>
        /// </list>
        /// </summary>
        public Boolean EvaluateFormulasBeforeSaving { get; set; } = false;

        /// <summary>
        /// Gets or sets the filter privacy flag. Set to null to leave the current property in saved workbook unchanged
        /// </summary>
        public Boolean? FilterPrivacy { get; set; }

        public Boolean GenerateCalculationChain { get; set; } = true;
        public Boolean ValidatePackage { get; set; }

        /// <summary>
        /// If set, forces the Open XML spreadsheet document type used during save
        /// (e.g. <see cref="SpreadsheetDocumentType.Workbook"/>, <see cref="SpreadsheetDocumentType.Template"/>)
        /// </summary>
        public SpreadsheetDocumentType? DocumentType { get; set; }
    }
}
