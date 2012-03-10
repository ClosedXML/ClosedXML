using System.IO;
using ClosedXML.Excel;
using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests
{
    internal static class TestHelper
    {
        //Note: Run example tests parameters
        public const string TestsOutputDirectory = @"D:\Excel Files\Tests\";
        public const string ActualTestResultPostFix = "";
        public static readonly string TestsExampleOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");
        
        private const bool CompareWithResources = true;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(null, ".Resource.Examples.");

        public static void SaveWorkbook(XLWorkbook workbook, string fileName)
        {
            workbook.SaveAs(Path.Combine(TestsOutputDirectory, fileName));
        }

        public static void RunTestExample<T>(string filePartName)
                where T : IXLExample, new()
        {
            var example = new T();
            string filePath1 = Path.Combine(TestsExampleOutputDirectory, filePartName);

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName= Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            filePath1 = Path.Combine(directory, "z" + fileName);
            var filePath2 = Path.Combine(directory, fileName);
            //Run test
            example.Create(filePath1);
            new XLWorkbook(filePath1).SaveAs(filePath2);
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
                    using (var streamActual = File.OpenRead(filePath2))
                    {
                        string message;
                        success = ExcelDocsComparer.Compare(streamActual, streamExpected, out message);
                        Assert.IsTrue(success, message);
                    }
                }
            }
            finally
            {
                //if (success && File.Exists(filePath)) File.Delete(filePath);
            }
#pragma warning restore 162
        }
    }
}