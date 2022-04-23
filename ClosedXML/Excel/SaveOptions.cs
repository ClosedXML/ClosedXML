// Keep this file CodeMaid organised and cleaned
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

        public bool ConsolidateConditionalFormatRanges { get; set; } = true;
        public bool ConsolidateDataValidationRanges { get; set; } = true;
        public bool EvaluateFormulasBeforeSaving { get; set; } = false;

        /// <summary>
        /// Gets or sets the filter privacy flag. Set to null to leave the current property in saved workbook unchanged
        /// </summary>
        public bool? FilterPrivacy { get; set; }

        public bool GenerateCalculationChain { get; set; } = true;
        public bool ValidatePackage { get; set; }
    }
}
