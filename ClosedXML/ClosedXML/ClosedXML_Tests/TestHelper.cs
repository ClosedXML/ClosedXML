using System.IO;
using ClosedXML.Excel;

namespace ClosedXML_Tests
{
    internal static class TestHelper
    {
        public const string TestsOutputDirectory = @"C:\ClosedXML\Tests\";
        public static void SaveWorkbook(XLWorkbook workbook, string fileName)
        {
            workbook.SaveAs(Path.Combine(TestsOutputDirectory, fileName));
        } 
    }
}