using System.IO;
using ClosedXML.Excel;


namespace ClosedXML.Examples
{
    public class TransposeRanges:IXLExample
    {
        public void Create(string filePath)
        {
            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new BasicTable().Create(tempFile);
                var workbook = new XLWorkbook(tempFile);

                var ws = workbook.Worksheet(1);

                var rngTable = ws.Range("B2:F6");

                rngTable.Transpose(XLTransposeOptions.MoveCells);

                ws.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
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