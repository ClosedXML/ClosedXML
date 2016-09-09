using System;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class LoadFiles
    {
        public static void LoadAllFiles()
        {
            var forLoadingFolder = @"C:\ClosedXML_Tests\Created";
            var forSavingFolder = @"C:\ClosedXML_Tests\Modified";

            foreach (var file in Directory.GetFiles(forLoadingFolder))
            {
                var fileInfo = new FileInfo(file);
                var fileName = fileInfo.Name;
                LoadAndSaveFile(forLoadingFolder + @"\" + fileName, forSavingFolder + @"\" + fileName);
            }

            //LoadAndSaveFile(forLoadingFolder + @"\StyleRowsColumns.xlsx", forSavingFolder + @"\StyleRowsColumns.xlsx");
        }

        private static void LoadAndSaveFile(String input, String output)
        {
            var wb = new XLWorkbook(input);
            wb.SaveAs(output);
            wb.SaveAs(output);
        }
    }
}
