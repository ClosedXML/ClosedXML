using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Threading;

namespace ClosedXML_Tests.Excel.Globalization
{
    [TestFixture]
    public class GlobalizationTests
    {
        [Test]
        [TestCase("A1*10", 1230d, "1230")]
        [TestCase("A1/10", 12.3, "12.3")]
        [TestCase("A1&\" cells\"", "123 cells", "123 cells")]
        [TestCase("A1&\"000\"", "123000", "123000")]
        [TestCase("ISNUMBER(A1)", true, "true")]
        [TestCase("ISBLANK(A1)", false, "false")]
        [TestCase("DATE(2018,1,28)", 43128d, "43128")]
        public void LoadFormulaCachedValue(string formula, object expectedValue, string expectedInnerText)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");

            using (var ms = new MemoryStream())
            {
                using (XLWorkbook book1 = new XLWorkbook())
                {
                    var sheet = book1.AddWorksheet("sheet1");
                    sheet.Cell("A1").Value = 123;
                    sheet.Cell("A2").FormulaA1 = formula;
                    var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

                    book1.SaveAs(ms, options);
                }
                ms.Position = 0;

                using (XLWorkbook book2 = new XLWorkbook(ms))
                {
                    var ws = book2.Worksheet(1);
                    var cellA2 = (XLCell)ws.Cell("A2");

                    Assert.That(cellA2.InnerText, Is.EqualTo(expectedInnerText));
                    Assert.That(cellA2.CachedValue, Is.EqualTo(expectedValue));
                    Assert.That(cellA2.NeedsRecalculation, Is.False);
                }
            }
        }
    }
}
