using System;
using System.IO;
using ClosedXML.Excel;

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
            var wb = new XLWorkbook(input);
            wb.SaveAs(output);
            wb.SaveAs(output);
        }
    }
}