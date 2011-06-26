using System.IO;
using ClosedXML.Excel;
using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests
{
    internal static class TestHelper
    {
        //Note: Run example tests parameters
        public const string TestsOutputDirectory = @"C:\Excel Files\";
        private const bool RemoveSuccessExampleFiles = false;
        private const bool CompareWithResources = false;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(null, ".Resources.");

        public static void SaveWorkbook(XLWorkbook workbook, string fileName)
        {
            workbook.SaveAs(Path.Combine(TestsOutputDirectory, fileName));
        }

        public static void RunTestExample<T>(string filePartName)
                where T : IXLExample, new()
        {
            var example = new T();
            string filePath = Path.Combine(TestsOutputDirectory, filePartName);

            //Run test
            example.Create(filePath);
            bool success = true;
#pragma warning disable 162
            try
            {
                //Compare
                // ReSharper disable ConditionIsAlwaysTrueOrFalse
                if (CompareWithResources)
                        // ReSharper restore ConditionIsAlwaysTrueOrFalse

                {
                    string resourcePath = filePartName.Replace('\\', '.').TrimStart('.');
                    using (var streamExpected = _extractor.ReadFileFromResToStream(resourcePath))
                    {
                        using (var streamActual = _extractor.ReadFileFromResToStream(resourcePath))
                        {
                            string message;
                            success = ExcelDocsComparer.Compare(streamActual, streamExpected, out message);
                            Assert.IsTrue(success, message);
                        }
                    }
                }
            }
            finally
            {
                // ReSharper disable ConditionIsAlwaysTrueOrFalse
                if (RemoveSuccessExampleFiles && success && File.Exists(filePath))
                        // ReSharper restore ConditionIsAlwaysTrueOrFalse
                {
                    File.Delete(filePath);
                }
            }
#pragma warning restore 162
        }
    }
}