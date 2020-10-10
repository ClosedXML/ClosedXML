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

        public Boolean ConsolidateConditionalFormatRanges { get; set; } = true;
        public Boolean ConsolidateDataValidationRanges { get; set; } = true;
        public Boolean EvaluateFormulasBeforeSaving { get; set; } = false;

        /// <summary>
        /// Gets or sets the filter privacy flag. Set to null to leave the current property in saved workbook unchanged
        /// </summary>
        public Boolean? FilterPrivacy { get; set; }

        public Boolean GenerateCalculationChain { get; set; } = true;
        public Boolean ValidatePackage { get; set; }
    }
}
