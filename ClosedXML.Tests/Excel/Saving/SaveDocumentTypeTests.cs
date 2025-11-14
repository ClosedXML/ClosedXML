using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests.Excel.Saving
{
    [TestFixture]
    public class SaveDocumentTypeTests
    {
        [Test]
        public void SaveAs_File_OverridesDocumentType()
        {
            using var template = new XLWorkbook();
            template.AddWorksheet("Sheet1");
            var tmpXltx = Path.ChangeExtension(Path.GetTempFileName(), ".xltx");
            var outXlsx = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");

            template.SaveAs(tmpXltx); // creates a Template file by extension
            using var wb = new XLWorkbook(tmpXltx);

            var options = new SaveOptions { DocumentType = SpreadsheetDocumentType.Workbook };
            wb.SaveAs(outXlsx, options);

            using var pkg = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(outXlsx, false);
            Assert.AreEqual(SpreadsheetDocumentType.Workbook, pkg.DocumentType);
        }

        [Test]
        public void SaveAs_Stream_UsesExplicitDocumentType()
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");

            using var ms = new MemoryStream();
            wb.SaveAs(ms, new SaveOptions { DocumentType = SpreadsheetDocumentType.Template });

            ms.Position = 0;
            using var pkg = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ms, false);
            Assert.AreEqual(SpreadsheetDocumentType.Template, pkg.DocumentType);
        }
    }
}
