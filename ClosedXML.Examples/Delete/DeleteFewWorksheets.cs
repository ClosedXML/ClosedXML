using ClosedXML.Excel;
using System.IO;

namespace ClosedXML.Examples.Delete
{
    public class DeleteFewWorksheets : IXLExample
    {
        public void Create(string filePath)
        {
            var tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                //Note: Prepare
                {
                    using var workbook = new XLWorkbook();
                    workbook.Worksheets.Add("1");
                    workbook.Worksheets.Add("2");
                    workbook.Worksheets.Add("3");
                    workbook.Worksheets.Add("4");
                    workbook.SaveAs(tempFile);
                }

                //Note: Delete few worksheet
                {
                    using var workbook = new XLWorkbook(tempFile);
                    workbook.Worksheets.Delete("1");
                    workbook.Worksheets.Delete("2");
                    workbook.SaveAs(filePath);
                }
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}