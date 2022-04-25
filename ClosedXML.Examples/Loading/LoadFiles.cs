using ClosedXML.Excel;
using System;
using System.IO;

namespace ClosedXML.Examples
{
    public class LoadFiles
    {
        public static void LoadAllFiles()
        {
            foreach (var file in Directory.GetFiles(Program.BaseCreatedDirectory))
            {
                var fileInfo = new FileInfo(file);
                var fileName = fileInfo.Name;
                LoadAndSaveFile(Path.Combine(Program.BaseCreatedDirectory, fileName), Path.Combine(Program.BaseModifiedDirectory, fileName));
            }
        }

        private static void LoadAndSaveFile(String input, String output)
        {
            using var wb = new XLWorkbook(input);
            wb.SaveAs(output);
            wb.SaveAs(output);
        }
    }
}