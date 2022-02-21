using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class FormulasWithEvaluation : Formulas
    {
        public override void Create(string filePath)
        {
            base.Create(filePath);
            using (var wb = new XLWorkbook(filePath))
            {
                wb.Save(true, true);
            }
        }
    }
}
