using System;
using System.IO;
using System.Threading;
using ClosedXML.Excel;
using ClosedXML_Examples;
using DocumentFormat.OpenXml.Drawing;
using NUnit.Framework;
using Path = System.IO.Path;

namespace ClosedXML_Tests
{
    internal static class TestHelper
    {
        public static string CurrencySymbol
        {
            get { return Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol; }
        }

        //Note: Run example tests parameters
        public static string TestsOutputDirectory
        {
            get { 
                return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location); 
            }

        }

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
            // Make sure tests run on a deterministic culture
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

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
                        var formattedMessage =
                            String.Format(
                                "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                                filePath2, resourcePath, message);

                        Assert.IsTrue(success, formattedMessage);
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