using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Formula
{
    public class XLFormulaDefinitionRepositoryTests
    {
        [Test]
        public void CanStoreFormulaDefinitions()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();

                var formula1 = new XLFormulaDefinition("=B1", (XLAddress) ws.Cell("A1").Address);
                var formula2 = new XLFormulaDefinition("=B2", (XLAddress) ws.Cell("A2").Address);

                var formula3 = wb.FormulaDefinitionRepository.Store(formula1);
                var formula4 = wb.FormulaDefinitionRepository.Store(formula2);

                Assert.AreSame(formula1, formula3);
                Assert.AreSame(formula1, formula4);
            }
        }

        [Test]
        public void FormulaDefinitionRepositoryCleanedOnDispose()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var formula1 = new XLFormulaDefinition("=B1", (XLAddress) ws.Cell("A1").Address);
            var formula2 = new XLFormulaDefinition("=B2", (XLAddress) ws.Cell("A2").Address);

            wb.FormulaDefinitionRepository.Store(formula1);

            wb.Dispose();

            var formula4 = wb.FormulaDefinitionRepository.Store(formula2);

            Assert.AreNotSame(formula1, formula4);
        }
    }
}
