using ClosedXML.Examples;
using ClosedXML.Excel;
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
        public static string CurrencySymbol
        {
            get { return Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol; }
        }

        //Note: Run example tests parameters
        public static string TestsOutputDirectory
        {
            get
            {
                return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Generated");
            }
        }

        public const string ActualTestResultPostFix = "";
        public static readonly string ExampleTestsOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");

        private const bool CompareWithResources = true;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(".Resource.");

        public static void SaveWorkbook(XLWorkbook workbook, params string[] fileNameParts)
        {
            workbook.SaveAs(Path.Combine(new string[] { TestsOutputDirectory }.Concat(fileNameParts).ToArray()), true);
        }

        // Because different fonts are installed on Unix,
        // the columns widths after AdjustToContents() will
        // cause the tests to fail.
        // Therefore we ignore the width attribute when running on Unix
        public static bool StripColumnWidths { get { return IsRunningOnUnix; } }

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
            string filePath1 = Path.Combine(new List<string>() { ExampleTestsOutputDirectory }.Concat(pathParts).ToArray());

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
                wb.SaveAs(filePath2, validate: true, evaluateFormulae);

            // Also load from template and save it again - but not necessary to test against reference file
            // We're just testing that it can save.
            using (var ms = new MemoryStream())
            using (var wb = XLWorkbook.OpenFromTemplate(filePath1))
                wb.SaveAs(ms, validate: true, evaluateFormulae);

            if (CompareWithResources)
            {
                string resourcePath = "Examples." + filePartName.Replace('\\', '.').TrimStart('.');
                using (var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
                using (var streamActual = File.OpenRead(filePath2))
                {
                    var success = ExcelDocsComparer.Compare(streamActual, streamExpected, out string message);
                    var formattedMessage =
                        String.Format(
                            "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                            filePath2, resourcePath, message);

                    Assert.IsTrue(success, formattedMessage);
                }
            }
        }

        /// <summary>
        /// Create a workbook and compare it with a saved resource.
        /// </summary>
        /// <param name="workbookGenerator">A function that gets an empty workbook and fills it with data.</param>
        /// <param name="referenceResource">Reference workbook saved in resources</param>
        /// <param name="evaluateFormulae">Should formulas of created workbook be evaluated and values saved?</param>
        /// <param name="validate">Should the created workbook be validated during by OpenXmlSdk validator?</param>
        public static void CreateAndCompare(Action<XLWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false, bool validate = true)
        {
            CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();
                workbookGenerator(wb);
                return wb;
            }, referenceResource, evaluateFormulae, validate);
        }

        public static void CreateAndCompare(Func<IXLWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false, bool validate = true)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string[] pathParts = referenceResource.Split(new char[] { '\\' });
            string filePath1 = Path.Combine(new List<string>() { TestsOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            var filePath2 = Path.Combine(directory, fileName);

            using (var wb = workbookGenerator.Invoke())
                wb.SaveAs(filePath2, validate, evaluateFormulae);

            if (CompareWithResources)
            {
                string resourcePath = referenceResource.Replace('\\', '.').TrimStart('.');
                using (var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
                using (var streamActual = File.OpenRead(filePath2))
                {
                    var success = ExcelDocsComparer.Compare(streamActual, streamExpected, out string message);
                    var formattedMessage =
                        String.Format(
                            "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                            filePath2, resourcePath, message);

                    Assert.IsTrue(success, formattedMessage);
                }
            }
        }

        /// <summary>
        /// Load a file from the <paramref name="loadResourcePath"/>, modify it, save it through ClosedXML
        /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
        /// </summary>
        /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
        public static void LoadModifyAndCompare(string loadResourcePath, Action<XLWorkbook> modify, string expectedOutputResourcePath, bool evaluateFormulae = false, bool validate = true)
        {
            using var stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
            using var ms = new MemoryStream();
            CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                modify(wb);
                wb.SaveAs(ms, validate);
                return wb;
            }, expectedOutputResourcePath, evaluateFormulae, validate);
        }

        /// <summary>
        /// Load a file from the <paramref name="loadResourcePath"/>, save it through ClosedXML without modifications
        /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
        /// </summary>
        /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
        public static void LoadSaveAndCompare(string loadResourcePath, string expectedOutputResourcePath, bool evaluateFormulae = false, bool validate = true)
        {
            LoadModifyAndCompare(loadResourcePath, _ => { }, expectedOutputResourcePath, evaluateFormulae, validate);
        }

        /// <summary>
        /// A testing method to load a workbook from resource and assert the state of the loaded workbook.
        /// </summary>
        public static void LoadAndAssert(Action<XLWorkbook> assertWorkbook, string loadResourcePath)
        {
            using var stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
            using var wb = new XLWorkbook(stream);

            assertWorkbook(wb);
        }

        /// <summary>
        /// A testing method to load a workbook with a single worksheet from resource and assert
        /// the state of the loaded workbook.
        /// </summary>
        public static void LoadAndAssert(Action<XLWorkbook, IXLWorksheet> assertWorksheet, string loadResourcePath)
        {
            LoadAndAssert(wb =>
            {
                var ws = wb.Worksheets.Single();
                assertWorksheet(wb, ws);
            }, loadResourcePath);
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
            using (var stream = GetStreamFromResource(GetResourcePath(filePartName)))
            {
                Assert.DoesNotThrow(() => wb = new XLWorkbook(stream), "Unable to load resource {0}", filePartName);
            }
        }

        public static IEnumerable<String> ListResourceFiles(Func<String, Boolean> predicate = null)
        {
            return _extractor.GetFileNames(predicate);
        }

        /// <summary>
        /// A method for testing of a saving and loading capabilities of ClosedXML. Use this
        /// method to check properties are correctly saved and loaded.
        /// </summary>
        /// <remarks>This method is specialized, so it only works on one sheet.</remarks>
        /// <param name="createWorksheet">
        /// Method to setup a worksheet that will be saved and the saved file will be compared to
        /// <paramref name="referenceResource"/>.
        /// </param>
        /// <param name="assertLoadedWorkbook">
        /// <paramref name="referenceResource"/> will be loaded and this method will check that it
        /// was loaded correctly (i.e. properties are what was set in <paramref name="createWorksheet"/>).
        /// </param>
        /// <param name="referenceResource">Saved reference file.</param>
        public static void CreateSaveLoadAssert(Action<XLWorkbook, IXLWorksheet> createWorksheet, Action<XLWorkbook, IXLWorksheet> assertLoadedWorkbook, string referenceResource)
        {
            CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();
                createWorksheet(wb, ws);
            }, referenceResource);
            LoadAndAssert(assertLoadedWorkbook, referenceResource);
        }

        /// <summary>
        /// Basically can survive through save and load cycle. Doesn't check against actual file.
        /// Useful for testing is internal structures are correctly initialized after load.
        /// </summary>
        /// <param name="createWorksheet">Code to create a workbook.</param>
        /// <param name="assertLoadedWorkbook">Method to assert that workbook was loaded correctly.</param>
        public static void CreateSaveLoadAssert(Action<XLWorkbook, IXLWorksheet> createWorksheet, Action<XLWorkbook, IXLWorksheet> assertLoadedWorkbook, bool validate = true, bool evaluateFormulas = false)
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                createWorksheet(wb, ws);
                wb.SaveAs(ms, validate, evaluateFormulas);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.Single();
                assertLoadedWorkbook(wb, ws);
            }
        }
    }
}
