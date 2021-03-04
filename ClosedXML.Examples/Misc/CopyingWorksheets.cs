using ClosedXML.Excel;
using ClosedXML.Examples.Tables;
using System.IO;

namespace ClosedXML.Examples.Misc
{
    public class CopyingWorksheets : IXLExample
    {
        public void Create(string filePath)
        {
            string tempFile1 = ExampleHelper.GetTempFilePath(filePath);
            string tempFile2 = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new UsingTables().Create(tempFile1);
                var wb = new XLWorkbook(tempFile1);

                var wsSource = wb.Worksheet(1);
                // Copy the worksheet to a new sheet in this workbook
                wsSource.CopyTo("Copy");

                // We're going to open another workbook to show that you can
                // copy a sheet from one workbook to another:
                new BasicTable().Create(tempFile2);
                var wbSource = new XLWorkbook(tempFile2);
                wbSource.Worksheet(1).CopyTo(wb, "Copy From Other");

                // Save the workbook with the 2 copies
                wb.SaveAs(filePath);
            }
            finally
            {
                if (File.Exists(tempFile1))
                {
                    File.Delete(tempFile1);
                }
                if (File.Exists(tempFile2))
                {
                    File.Delete(tempFile2);
                }
            }
        }
    }
}
