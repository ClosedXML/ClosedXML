using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;

namespace ClosedXML_Examples
{
    public class LoadFiles
    {
        public static void LoadAllFiles()
        {
            var forLoadingFolder = @"C:\Excel Files\Created";
            var forSavingFolder = @"C:\Excel Files\Modified";

            foreach (var file in Directory.GetFiles(forLoadingFolder))
            {
                var fileInfo = new FileInfo(file);
                var fileName = fileInfo.Name;
                LoadAndSaveFile(forLoadingFolder + @"\" + fileName, forSavingFolder + @"\" + fileName);
            }

            //LoadAndSaveFile(forLoadingFolder + @"\StyleWorksheet.xlsx", forSavingFolder + @"\StyleWorksheet.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "DataTypes.xlsx", forSavingFolder + "DataTypes.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "MultipleSheets.xlsx", forSavingFolder + "MultipleSheets.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "styleNumberFormat.xlsx", forSavingFolder + "styleNumberFormat.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "styleFill.xlsx", forSavingFolder + "styleFill.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "styleAlignment.xlsx", forSavingFolder + "styleAlignment.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "styleBorder.xlsx", forSavingFolder + "styleBorder.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "styleFont.xlsx", forSavingFolder + "styleFont.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "MergedCells.xlsx", forSavingFolder + "MergedCells.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "ColumnSettings.xlsx", forSavingFolder + "ColumnSettings.xlsx");
            //LoadAndSaveFile(forLoadingFolder + "RowSettings.xlsx", forSavingFolder + "RowSettings.xlsx");
        }

        private static void LoadAndSaveFile(String input, String output)
        {
            var wb = new XLWorkbook(input);
            wb.SaveAs(output);
        }
    }
}
