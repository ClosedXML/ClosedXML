using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Examples.Delete
{
    public class DeleteFewWorksheets:IXLExample
    {
        public void Create(string filePath)
        {
            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                //Note: Prepare
                {
                    var workbook = new XLWorkbook();
                    workbook.Worksheets.Add("1");
                    workbook.Worksheets.Add("2");
                    workbook.Worksheets.Add("3");
                    workbook.Worksheets.Add("4");
                    workbook.SaveAs(tempFile);
                }

                //Note: Delate few worksheet
                {
                    var workbook = new XLWorkbook(tempFile);
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