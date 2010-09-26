using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class LoadFiles
    {
        public static void LoadAllFiles()
        {
            var forLoadingFolder = @"C:\Excel Files\ForLoading\";
            var forSavingFolder = @"C:\Excel Files\Modified\";

            LoadAndSaveFile(forLoadingFolder + "HelloWorld.xlsx", forSavingFolder + "HelloWorld.xlsx");
            LoadAndSaveFile(forLoadingFolder + "DataTypes.xlsx", forSavingFolder + "DataTypes.xlsx");
            LoadAndSaveFile(forLoadingFolder + "MultipleSheets.xlsx", forSavingFolder + "MultipleSheets.xlsx");
            LoadAndSaveFile(forLoadingFolder + "styleNumberFormat.xlsx", forSavingFolder + "styleNumberFormat.xlsx");
            LoadAndSaveFile(forLoadingFolder + "styleFill.xlsx", forSavingFolder + "styleFill.xlsx");
            LoadAndSaveFile(forLoadingFolder + "styleAlignment.xlsx", forSavingFolder + "styleAlignment.xlsx");
        }

        private static void LoadAndSaveFile(String input, String output)
        {
            var wb = new XLWorkbook();
            wb.Load(input);
            wb.SaveAs(output);
        }
    }
}
