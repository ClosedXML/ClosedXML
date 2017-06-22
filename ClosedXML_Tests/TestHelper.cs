using ClosedXML.Excel;
using ClosedXML_Examples;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
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
            get
            {
                return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            }
        }

        public const string ActualTestResultPostFix = "";
        public static readonly string TestsExampleOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");

        private const bool CompareWithResources = true;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(null, ".Resource.Examples.");

        public static void SaveWorkbook(XLWorkbook workbook, params string[] fileNameParts)
        {
            workbook.SaveAs(Path.Combine(new string[] { TestsOutputDirectory }.Concat(fileNameParts).ToArray()), true);
        }

        // Because different fonts are installed on Unix,
        // the columns widths after AdjustToContents() will
        // cause the tests to fail.
        // Therefore we ignore the width attribute when running on Unix
        public static bool IsRunningOnUnix
        {
            get
            {
                int p = (int)Environment.OSVersion.Platform;
                return ((p == 4) || (p == 6) || (p == 128));
            }
        }

        public static void RunTestExample<T>(string filePartName, bool evaluateFormulae = false)
                where T : IXLExample, new()
        {
            // Make sure tests run on a deterministic culture
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            var example = new T();
            string[] pathParts = filePartName.Split(new char[] { '\\' });
            string filePath1 = Path.Combine(new List<string>() { TestsExampleOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            filePath1 = Path.Combine(directory, "z" + fileName);
            var filePath2 = Path.Combine(directory, fileName);
            //Run test
            example.Create(filePath1);
            using (var wb = new XLWorkbook(filePath1))
                wb.SaveAs(filePath2, true, evaluateFormulae);

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
                        success = ExcelDocsComparer.Compare(streamActual, streamExpected, TestHelper.IsRunningOnUnix, out message);
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

        public static string GetResourcePath(string filePartName)
        {
            return filePartName.Replace('\\', '.').TrimStart('.');
        }

        public static Stream GetStreamFromResource(string resourcePath)
        {
            var extractor = new ResourceFileExtractor(null, ".Resource.");
            return extractor.ReadFileFromResToStream(resourcePath);
        }

        public static void LoadFile(string filePartName)
        {
            using (var stream = GetStreamFromResource(GetResourcePath(filePartName)))
            {
                var wb = new XLWorkbook(stream);
                wb.Dispose();
            }
        }
    }
}
