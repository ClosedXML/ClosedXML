using System;

namespace ClosedXML.Excel
{
    public sealed class SaveOptions
    {
        public SaveOptions()
        {
#if DEBUG
            this.ValidatePackage = true;
#else
            this.ValidatePackage = false;
#endif

            this.EvaluateFormulasBeforeSaving = false;
            this.GenerateCalculationChain = true;
        }

        public Boolean ValidatePackage;
        public Boolean EvaluateFormulasBeforeSaving;
        public Boolean GenerateCalculationChain;
    }
}
