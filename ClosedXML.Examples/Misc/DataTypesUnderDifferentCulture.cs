using System;
using ClosedXML.Excel;
using System.Threading;
using System.Globalization;
using System.IO;

namespace ClosedXML.Examples.Misc
{
    public class DataTypesUnderDifferentCulture : IXLExample
    {
        public void Create(string filePath)
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            // Set thread culture to French, which should format numbers using decimal COMMA
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("fr-FR");

            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new DataTypes().Create(tempFile);
                var workbook = new XLWorkbook(tempFile);
                workbook.SaveAs(filePath);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}
