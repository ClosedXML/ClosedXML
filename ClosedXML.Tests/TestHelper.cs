using ClosedXML.Examples;
using ClosedXML.Excel;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Path = System.IO.Path;

namespace ClosedXML.Tests
{
    internal static class TestHelper
    {
        public static string CurrencySymbol => Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol;

        public static string TestsOutputDirectory => Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Generated");

        public const string ActualTestResultPostFix = "";
        public static readonly string ExampleTestsOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");

        private const bool CompareWithResources = true;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(".Resource.");

        public static void SaveWorkbook(XLWorkbook workbook, params string[] fileNameParts)
        {
            workbook.SaveAs(Path.Combine(new string[] { TestsOutputDirectory }.Concat(fileNameParts).ToArray()), true);
        }

        public static void RunTestExample<T>(string filePartName, bool evaluateFormula = false, string expectedDiff = null, bool ignoreColumnFormats = false)
                where T : IXLExample, new()
        {
            // Make sure tests run on a deterministic culture
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            var example = new T();
            var pathParts = filePartName.Split(new[] { '\\' });
            var filePath1 = Path.Combine(new List<string> { ExampleTestsOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            filePath1 = Path.Combine(directory, "z" + fileName);
            var filePath2 = Path.Combine(directory, fileName);

            //Run test
            example.Create(filePath1);
            example.Create(filePath1);
            using (var wb = new XLWorkbook(filePath1))
            {
                wb.SaveAs(filePath2, validate: true, evaluateFormula);
            }

            // Also load from template and save it again - but not necessary to test against reference file
            // We're just testing that it can save.
            using (var ms = new MemoryStream())
            using (var wb = XLWorkbook.OpenFromTemplate(filePath1))
            {
                wb.SaveAs(ms, validate: true, evaluateFormula);
            }

            // Uncomment to replace expectation running .net6.0,
            //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "resource", "examples", filePartName.Replace("\\", "/")));
            //File.Copy(filePath2, expectedFileInVsSolution, true);
            // Uncomment to replace expectation running .net4.8,
            //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("ClosedXML.Tests", "Resource", "Examples", filePartName.Replace("\\", "/")));
            //File.Copy(filePath2, expectedFileInVsSolution, true);

            if (CompareWithResources)
            {
                var resourcePath = "Examples." + filePartName.Replace('\\', '.').TrimStart('.');
                using var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath);
                using var streamActual = File.OpenRead(filePath2);
                var success = ExcelDocsComparer.Compare(streamActual, streamExpected, out var message, ignoreColumnFormats);
                var formattedMessage = $"Actual file '{filePath2}' is different than the expected file '{resourcePath}'. The difference is: '{message}'.";

                if (success)
                {
                    return;
                }

                SaveToTestresults(streamExpected, "Expected" + resourcePath);
                SaveToTestresults(streamActual, "Actual" + resourcePath);

                if (string.IsNullOrEmpty(expectedDiff))
                {
                    Assert.Fail(formattedMessage);
                }

                Assert.That(message, Is.EqualTo(expectedDiff), $"Actual diff '{message}' differs to expected diff '{expectedDiff}', file '{resourcePath}'");
            }
        }

        private static void SaveToTestresults(Stream streamExpected, string filename)
        {
            var testResultDirectory = Path.Combine(TestsOutputDirectory, "../../../../../TestResult");
            if (!Directory.Exists(testResultDirectory))
            {
                Directory.CreateDirectory(testResultDirectory);
            }
            streamExpected.Position = 0;
            var path = Path.Combine(testResultDirectory, filename);

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using var expectedFile = new FileStream(path, FileMode.Create);
            streamExpected.CopyTo(expectedFile);
        }

        public static void CreateAndCompare(Func<IXLWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false, bool ignoreColumnFormats = false)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            var pathParts = referenceResource.Split(new[] { '\\' });
            var filePath1 = Path.Combine(new List<string> { TestsOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            var filePath2 = Path.Combine(directory, fileName);

            using (var wb = workbookGenerator.Invoke())
            {
                wb.SaveAs(filePath2, true, evaluateFormulae);
            }

            if (CompareWithResources)
            {
                var resourcePath = referenceResource.Replace('\\', '.').TrimStart('.');
                using var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath);
                using var streamActual = File.OpenRead(filePath2);
                var success = ExcelDocsComparer.Compare(streamActual, streamExpected, out var message, ignoreColumnFormats);
                var formattedMessage =
                    string.Format(
                        "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                        filePath2, resourcePath, message);

                Assert.IsTrue(success, formattedMessage);
            }
        }

        public static string GetResourcePath(string filePartName)
        {
            return filePartName.Replace('\\', '.').TrimStart('.');
        }

        public static Stream GetStreamFromResource(string resourcePath)
        {
            return _extractor.ReadFileFromResourceToStream(resourcePath);
        }

        public static void LoadFile(string filePartName)
        {
            IXLWorkbook wb;
            using var stream = GetStreamFromResource(GetResourcePath(filePartName));
            Assert.DoesNotThrow(() => wb = new XLWorkbook(stream), "Unable to load resource {0}", filePartName);
        }

        public static IEnumerable<string> ListResourceFiles(Func<string, bool> predicate = null)
        {
            return _extractor.GetFileNames(predicate);
        }
    }
}