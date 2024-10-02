#nullable disable

using ClosedXML.Excel.CalcEngine;
using ClosedXML.Graphics;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    public enum XLCalculateMode
    {
        Auto,
        AutoNoTable,
        Manual,
        Default
    };

    public enum XLReferenceStyle
    {
        R1C1,
        A1,
        Default
    };

    public enum XLCellSetValueBehavior
    {
        /// <summary>
        ///   Analyze input string and convert value. For avoid analyzing use escape symbol '
        /// </summary>
        Smart = 0,

        /// <summary>
        ///   Direct set value. If value has unsupported type - value will be stored as string returned by <see
        ///    cref = "object.ToString()" />
        /// </summary>
        Simple = 1,
    }

    public partial class XLWorkbook : IXLWorkbook
    {
        #region Static

        public static IXLStyle DefaultStyle
        {
            get
            {
                return XLStyle.Default;
            }
        }

        internal static XLStyleValue DefaultStyleValue
        {
            get
            {
                return XLStyleValue.Default;
            }
        }

        public static Double DefaultRowHeight { get; private set; }
        public static Double DefaultColumnWidth { get; private set; }

        public static IXLPageSetup DefaultPageOptions
        {
            get
            {
                var defaultPageOptions = new XLPageSetup(null, null)
                {
                    PageOrientation = XLPageOrientation.Default,
                    Scale = 100,
                    PaperSize = XLPaperSize.LetterPaper,
                    Margins = new XLMargins
                    {
                        Top = 0.75,
                        Bottom = 0.5,
                        Left = 0.75,
                        Right = 0.75,
                        Header = 0.5,
                        Footer = 0.75
                    },
                    ScaleHFWithDocument = true,
                    AlignHFWithMargins = true,
                    PrintErrorValue = XLPrintErrorValues.Displayed,
                    ShowComments = XLShowCommentsValues.None
                };
                return defaultPageOptions;
            }
        }

        public static IXLOutline DefaultOutline
        {
            get
            {
                return new XLOutline(null)
                {
                    SummaryHLocation = XLOutlineSummaryHLocation.Right,
                    SummaryVLocation = XLOutlineSummaryVLocation.Bottom
                };
            }
        }

        /// <summary>
        ///   Behavior for <see cref = "IXLCell.set_Value" />
        /// </summary>
        public static XLCellSetValueBehavior CellSetValueBehavior { get; set; }

        public static XLWorkbook OpenFromTemplate(String path)
        {
            return new XLWorkbook(path, asTemplate: true);
        }

        #endregion Static

        internal readonly List<UnsupportedSheet> UnsupportedSheets =
            new List<UnsupportedSheet>();

        internal IXLGraphicEngine GraphicEngine { get; }

        internal double DpiX { get; }

        internal double DpiY { get; }

        internal XLPivotCaches PivotCachesInternal { get; }

        internal SharedStringTable SharedStringTable { get; } = new();

        #region Nested Type : XLLoadSource

        private enum XLLoadSource
        {
            New,
            File,
            Stream
        };

        #endregion Nested Type : XLLoadSource

        internal XLWorksheets WorksheetsInternal { get; private set; }

        /// <summary>
        ///   Gets an object to manipulate the worksheets.
        /// </summary>
        public IXLWorksheets Worksheets
        {
            get { return WorksheetsInternal; }
        }

        internal XLDefinedNames DefinedNamesInternal { get; }

        [Obsolete($"Use {nameof(DefinedNames)} instead.")]
        public IXLDefinedNames NamedRanges => DefinedNamesInternal;

        /// <summary>
        ///   Gets an object to manipulate this workbook's named ranges.
        /// </summary>
        public IXLDefinedNames DefinedNames => DefinedNamesInternal;

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        public IXLTheme Theme { get; private set; }

        /// <summary>
        /// All pivot caches in the workbook, whether they have a pivot table or not.
        /// </summary>
        public IXLPivotCaches PivotCaches => PivotCachesInternal;

        /// <summary>
        ///   Gets or sets the default style for the workbook.
        ///   <para>All new worksheets will use this style.</para>
        /// </summary>
        public IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets or sets the default row height for the workbook.
        ///   <para>All new worksheets will use this row height.</para>
        /// </summary>
        public Double RowHeight { get; set; }

        /// <summary>
        ///   Gets or sets the default column width for the workbook.
        ///   <para>All new worksheets will use this column width.</para>
        /// </summary>
        public Double ColumnWidth { get; set; }

        /// <summary>
        ///   Gets or sets the default page options for the workbook.
        ///   <para>All new worksheets will use these page options.</para>
        /// </summary>
        public IXLPageSetup PageOptions { get; set; }

        /// <summary>
        ///   Gets or sets the default outline options for the workbook.
        ///   <para>All new worksheets will use these outline options.</para>
        /// </summary>
        public IXLOutline Outline { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's properties.
        /// </summary>
        public XLWorkbookProperties Properties { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's calculation mode.
        /// </summary>
        public XLCalculateMode CalculateMode { get; set; }

        public Boolean CalculationOnSave { get; set; }
        public Boolean ForceFullCalculation { get; set; }
        public Boolean FullCalculationOnLoad { get; set; }
        public Boolean FullPrecision { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's reference style.
        /// </summary>
        public XLReferenceStyle ReferenceStyle { get; set; }

        public IXLCustomProperties CustomProperties { get; private set; }

        public Boolean ShowFormulas { get; set; }
        public Boolean ShowGridLines { get; set; }
        public Boolean ShowOutlineSymbols { get; set; }
        public Boolean ShowRowColHeaders { get; set; }
        public Boolean ShowRuler { get; set; }
        public Boolean ShowWhiteSpace { get; set; }
        public Boolean ShowZeros { get; set; }
        public Boolean RightToLeft { get; set; }

        public Boolean DefaultShowFormulas
        {
            get { return false; }
        }

        public Boolean DefaultShowGridLines
        {
            get { return true; }
        }

        public Boolean DefaultShowOutlineSymbols
        {
            get { return true; }
        }

        public Boolean DefaultShowRowColHeaders
        {
            get { return true; }
        }

        public Boolean DefaultShowRuler
        {
            get { return true; }
        }

        public Boolean DefaultShowWhiteSpace
        {
            get { return true; }
        }

        public Boolean DefaultShowZeros
        {
            get { return true; }
        }

        public IXLFileSharing FileSharing { get; } = new XLFileSharing();

        public Boolean DefaultRightToLeft
        {
            get { return false; }
        }

        private void InitializeTheme()
        {
            Theme = new XLTheme
            {
                Text1 = XLColor.FromHtml("#FF000000"),
                Background1 = XLColor.FromHtml("#FFFFFFFF"),
                Text2 = XLColor.FromHtml("#FF1F497D"),
                Background2 = XLColor.FromHtml("#FFEEECE1"),
                Accent1 = XLColor.FromHtml("#FF4F81BD"),
                Accent2 = XLColor.FromHtml("#FFC0504D"),
                Accent3 = XLColor.FromHtml("#FF9BBB59"),
                Accent4 = XLColor.FromHtml("#FF8064A2"),
                Accent5 = XLColor.FromHtml("#FF4BACC6"),
                Accent6 = XLColor.FromHtml("#FFF79646"),
                Hyperlink = XLColor.FromHtml("#FF0000FF"),
                FollowedHyperlink = XLColor.FromHtml("#FF800080")
            };
        }

#nullable enable
        [Obsolete($"Use {nameof(DefinedName)} instead.")]
        public IXLDefinedName? NamedRange(String name) => DefinedName(name);

        /// <inheritdoc/>
        public IXLDefinedName? DefinedName(String name)
        {
            if (name.Contains("!"))
            {
                var split = name.Split('!');
                var first = split[0];
                var wsName = first.StartsWith("'") ? first.Substring(1, first.Length - 2) : first;
                var sheetlessName = split[1];
                if (TryGetWorksheet(wsName, out XLWorksheet ws))
                {
                    if (ws.DefinedNames.TryGetScopedValue(sheetlessName, out var sheetDefinedName))
                        return sheetDefinedName;
                }

                name = sheetlessName;
            }

            return DefinedNamesInternal.TryGetScopedValue(name, out var definedName) ? definedName : null;
        }
#nullable disable

        public Boolean TryGetWorksheet(String name, out IXLWorksheet worksheet)
        {
            if (TryGetWorksheet(name, out XLWorksheet foundSheet))
            {
                worksheet = foundSheet;
                return true;
            }

            worksheet = default;
            return false;
        }

        internal Boolean TryGetWorksheet(String name, [NotNullWhen(true)] out XLWorksheet worksheet)
        {
            return WorksheetsInternal.TryGetWorksheet(name, out worksheet);
        }

        public IXLRange RangeFromFullAddress(String rangeAddress, out IXLWorksheet ws)
        {
            if (!rangeAddress.Contains('!'))
            {
                ws = null;
                return null;
            }

            var split = rangeAddress.Split('!');
            var wsName = split[0].UnescapeSheetName();
            if (TryGetWorksheet(wsName, out XLWorksheet sheet))
            {
                ws = sheet;
                return sheet.Range(split[1]);
            }

            ws = null;
            return null;
        }

        public IXLCell CellFromFullAddress(String cellAddress, out IXLWorksheet ws)
        {
            if (!cellAddress.Contains('!'))
            {
                ws = null;
                return null;
            }

            var split = cellAddress.Split('!');
            var wsName = split[0].UnescapeSheetName();
            if (TryGetWorksheet(wsName, out XLWorksheet sheet))
            {
                ws = sheet;
                return sheet.Cell(split[1]);
            }

            ws = null;
            return null;
        }

        /// <summary>
        ///   Saves the current workbook.
        /// </summary>
        public void Save()
        {
#if DEBUG
            Save(true, false);
#else
            Save(false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook and optionally performs validation
        /// </summary>
        public void Save(Boolean validate, Boolean evaluateFormulae = false)
        {
            Save(new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void Save(SaveOptions options)
        {
            checkForWorksheetsPresent();
            if (_loadSource == XLLoadSource.New)
                throw new InvalidOperationException("This is a new file. Please use one of the 'SaveAs' methods.");

            if (_loadSource == XLLoadSource.Stream)
            {
                CreatePackage(_originalStream, false, _spreadsheetDocumentType, options);
            }
            else
                CreatePackage(_originalFile, _spreadsheetDocumentType, options);
        }

        /// <summary>
        ///   Saves the current workbook to a file.
        /// </summary>
        public void SaveAs(String file)
        {
#if DEBUG
            SaveAs(file, true, false);
#else
            SaveAs(file, false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook to a file and optionally validates it.
        /// </summary>
        public void SaveAs(String file, Boolean validate, Boolean evaluateFormulae = false)
        {
            SaveAs(file, new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void SaveAs(String file, SaveOptions options)
        {
            checkForWorksheetsPresent();

            var directoryName = Path.GetDirectoryName(file);
            if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

            if (_loadSource == XLLoadSource.New)
            {
                if (File.Exists(file))
                    File.Delete(file);

                CreatePackage(file, GetSpreadsheetDocumentType(file), options);
            }
            else if (_loadSource == XLLoadSource.File)
            {
                if (String.Compare(_originalFile.Trim(), file.Trim(), true) != 0)
                {
                    File.Copy(_originalFile, file, true);
                    File.SetAttributes(file, FileAttributes.Normal);
                }

                CreatePackage(file, GetSpreadsheetDocumentType(file), options);
            }
            else if (_loadSource == XLLoadSource.Stream)
            {
                _originalStream.Position = 0;

                using (var fileStream = File.Create(file))
                {
                    CopyStream(_originalStream, fileStream);
                    CreatePackage(fileStream, false, _spreadsheetDocumentType, options);
                }
            }

            _loadSource = XLLoadSource.File;
            _originalFile = file;
            _originalStream = null;
        }

        private static SpreadsheetDocumentType GetSpreadsheetDocumentType(string filePath)
        {
            var extension = Path.GetExtension(filePath);

            if (string.IsNullOrEmpty(extension)) throw new ArgumentException("Empty extension is not supported.");
            extension = extension.Substring(1).ToLowerInvariant();

            switch (extension)
            {
                case "xlsm":
                    return SpreadsheetDocumentType.MacroEnabledWorkbook;

                case "xltm":
                    return SpreadsheetDocumentType.MacroEnabledTemplate;

                case "xlsx":
                    return SpreadsheetDocumentType.Workbook;

                case "xltx":
                    return SpreadsheetDocumentType.Template;

                default:
                    throw new ArgumentException(String.Format("Extension '{0}' is not supported. Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.", extension));
            }
        }

        private void checkForWorksheetsPresent()
        {
            if (!Worksheets.Any())
                throw new InvalidOperationException("Workbooks need at least one worksheet.");
        }

        /// <summary>
        ///   Saves the current workbook to a stream.
        /// </summary>
        public void SaveAs(Stream stream)
        {
#if DEBUG
            SaveAs(stream, true, false);
#else
            SaveAs(stream, false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook to a stream and optionally validates it.
        /// </summary>
        public void SaveAs(Stream stream, Boolean validate, Boolean evaluateFormulae = false)
        {
            SaveAs(stream, new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void SaveAs(Stream stream, SaveOptions options)
        {
            checkForWorksheetsPresent();
            if (_loadSource == XLLoadSource.New)
            {
                // dm 20130422, this method or better the method SpreadsheetDocument.Create which is called
                // inside of 'CreatePackage' need a stream which CanSeek & CanRead
                // and an ordinary Response stream of a webserver can't do this
                // so we have to ask and provide a way around this
                if (stream.CanRead && stream.CanSeek && stream.CanWrite)
                {
                    // all is fine the package can be created in a direct way
                    CreatePackage(stream, true, _spreadsheetDocumentType, options);
                }
                else
                {
                    // the harder way
                    using (var ms = new MemoryStream())
                    {
                        CreatePackage(ms, true, _spreadsheetDocumentType, options);
                        // not really necessary, because I changed CopyStream too.
                        // but for better understanding and if somebody in the future
                        // provide an changed version of CopyStream
                        ms.Position = 0;
                        CopyStream(ms, stream);
                    }
                }
            }
            else if (_loadSource == XLLoadSource.File)
            {
                using (var fileStream = new FileStream(_originalFile, FileMode.Open, FileAccess.Read))
                {
                    CopyStream(fileStream, stream);
                }
                CreatePackage(stream, false, _spreadsheetDocumentType, options);
            }
            else if (_loadSource == XLLoadSource.Stream)
            {
                _originalStream.Position = 0;
                if (_originalStream != stream)
                    CopyStream(_originalStream, stream);

                CreatePackage(stream, false, _spreadsheetDocumentType, options);
            }

            _loadSource = XLLoadSource.Stream;
            _originalStream = stream;
            _originalFile = null;
        }

        internal static void CopyStream(Stream input, Stream output)
        {
            var buffer = new byte[8 * 1024];
            int len;
            // dm 20130422, it is always a good idea to rewind the input stream, or not?
            if (input.CanSeek)
                input.Seek(0, SeekOrigin.Begin);
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
                output.Write(buffer, 0, len);
            // dm 20130422, and flushing the output after write
            output.Flush();
        }

        public IXLTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
        {
            if (!TryGetTable(tableName, out var table, comparisonType))
                throw new ArgumentOutOfRangeException($"Table {tableName} was not found.");

            return table;
        }

        /// <summary>
        /// Try to find a table with <paramref name="tableName"/> in a workbook.
        /// </summary>
        internal bool TryGetTable(string tableName, out XLTable table, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
        {
            table = WorksheetsInternal
                .SelectMany<XLWorksheet, XLTable>(ws => ws.Tables)
                .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

            return table is not null;
        }

        /// <summary>
        /// Try to find a table that covers same area as the <paramref name="area"/> in a workbook.
        /// </summary>
        internal bool TryGetTable(XLBookArea area, out XLTable foundTable)
        {
            foreach (var sheet in WorksheetsInternal)
            {
                if (XLHelper.SheetComparer.Equals(sheet.Name, area.Name))
                {
                    foreach (var table in sheet.Tables)
                    {
                        if (table.Area != area.Area)
                            continue;

                        foundTable = table;
                        return true;
                    }

                    // No other sheet has correct name.
                    break;
                }
            }

            foundTable = null;
            return false;
        }

        public IXLWorksheet Worksheet(String name)
        {
            return WorksheetsInternal.Worksheet(name);
        }

        public IXLWorksheet Worksheet(Int32 position)
        {
            return WorksheetsInternal.Worksheet(position);
        }

        public IXLCustomProperty CustomProperty(String name)
        {
            return CustomProperties.CustomProperty(name);
        }

        public IXLCells FindCells(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (XLWorksheet ws in WorksheetsInternal)
            {
                foreach (XLCell cell in ws.CellsUsed(XLCellsUsedOptions.All))
                {
                    if (predicate(cell))
                        cells.Add(cell);
                }
            }
            return cells;
        }

        public IXLRows FindRows(Func<IXLRow, Boolean> predicate)
        {
            var rows = new XLRows(worksheet: null);
            foreach (XLWorksheet ws in WorksheetsInternal)
            {
                foreach (IXLRow row in ws.Rows().Where(predicate))
                    rows.Add(row as XLRow);
            }
            return rows;
        }

        public IXLColumns FindColumns(Func<IXLColumn, Boolean> predicate)
        {
            var columns = new XLColumns(worksheet: null);
            foreach (XLWorksheet ws in WorksheetsInternal)
            {
                foreach (IXLColumn column in ws.Columns().Where(predicate))
                    columns.Add(column as XLColumn);
            }
            return columns;
        }

        /// <summary>
        /// Searches the cells' contents for a given piece of text
        /// </summary>
        /// <param name="searchText">The search text.</param>
        /// <param name="compareOptions">The compare options.</param>
        /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
        public IEnumerable<IXLCell> Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false)
        {
            foreach (var ws in WorksheetsInternal)
            {
                foreach (var cell in ws.Search(searchText, compareOptions, searchFormulae))
                    yield return cell;
            }
        }

        #region Fields

        private XLLoadSource _loadSource = XLLoadSource.New;
        private String _originalFile;
        private Stream _originalStream;
        private XLWorkbookProtection _workbookProtection;

        #endregion Fields

        #region Constructor

        /// <summary>
        ///   Creates a new Excel workbook.
        /// </summary>
        public XLWorkbook()
            : this(new LoadOptions())
        {
        }

        internal XLWorkbook(String file, Boolean asTemplate)
            : this(new LoadOptions())
        {
            LoadSheetsFromTemplate(file, new LoadOptions());
        }

        /// <summary>
        ///   Opens an existing workbook from a file.
        /// </summary>
        /// <param name = "file">The file to open.</param>
        public XLWorkbook(String file)
            : this(file, new LoadOptions())
        {
        }

        public XLWorkbook(String file, LoadOptions loadOptions)
            : this(loadOptions)
        {
            _loadSource = XLLoadSource.File;
            _originalFile = file;
            _spreadsheetDocumentType = GetSpreadsheetDocumentType(_originalFile);
            Load(file, loadOptions);

            if (loadOptions.RecalculateAllFormulas)
                this.RecalculateAllFormulas();
        }

        /// <summary>
        ///   Opens an existing workbook from a stream.
        /// </summary>
        /// <param name = "stream">The stream to open.</param>
        public XLWorkbook(Stream stream)
            : this(stream, new LoadOptions())
        {
        }

        public XLWorkbook(Stream stream, LoadOptions loadOptions)
            : this(loadOptions)
        {
            _loadSource = XLLoadSource.Stream;
            _originalStream = stream;
            Load(stream, loadOptions);

            if (loadOptions.RecalculateAllFormulas)
                this.RecalculateAllFormulas();
        }

        public XLWorkbook(LoadOptions loadOptions)
        {
            if (loadOptions is null)
                throw new ArgumentNullException(nameof(loadOptions));

            DpiX = loadOptions.Dpi.X;
            DpiY = loadOptions.Dpi.Y;
            GraphicEngine = loadOptions.GraphicEngine ?? LoadOptions.DefaultGraphicEngine ?? DefaultGraphicEngine.Instance.Value;
            Protection = new XLWorkbookProtection(DefaultProtectionAlgorithm);
            DefaultRowHeight = 15;
            DefaultColumnWidth = 8.43;
            Style = new XLStyle(null, DefaultStyle);
            RowHeight = DefaultRowHeight;
            ColumnWidth = DefaultColumnWidth;
            PageOptions = DefaultPageOptions;
            Outline = DefaultOutline;
            Properties = new XLWorkbookProperties();
            CalculateMode = XLCalculateMode.Default;
            ReferenceStyle = XLReferenceStyle.Default;
            InitializeTheme();
            ShowFormulas = DefaultShowFormulas;
            ShowGridLines = DefaultShowGridLines;
            ShowOutlineSymbols = DefaultShowOutlineSymbols;
            ShowRowColHeaders = DefaultShowRowColHeaders;
            ShowRuler = DefaultShowRuler;
            ShowWhiteSpace = DefaultShowWhiteSpace;
            ShowZeros = DefaultShowZeros;
            RightToLeft = DefaultRightToLeft;
            WorksheetsInternal = new XLWorksheets(this);
            DefinedNamesInternal = new XLDefinedNames(this);
            PivotCachesInternal = new XLPivotCaches(this);
            CustomProperties = new XLCustomProperties(this);
            ShapeIdManager = new XLIdManager();
            Author = Environment.UserName;
        }

        #endregion Constructor

        #region Nested type: UnsupportedSheet

        internal sealed class UnsupportedSheet
        {
            public Boolean IsActive;
            public UInt32 SheetId;
            public Int32 Position;
        }

        #endregion Nested type: UnsupportedSheet

        public IXLCell Cell(String namedCell)
        {
            var namedRange = DefinedName(namedCell);
            if (namedRange != null)
            {
                return namedRange.Ranges?.FirstOrDefault()?.FirstCell();
            }
            else
                return CellFromFullAddress(namedCell, out _);
        }

        public IXLCells Cells(String namedCells)
        {
            return Ranges(namedCells).Cells();
        }

        public IXLRange Range(String range)
        {
            var namedRange = DefinedName(range);
            if (namedRange != null)
                return namedRange.Ranges.FirstOrDefault();
            else
                return RangeFromFullAddress(range, out _);
        }

        public IXLRanges Ranges(String ranges)
        {
            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (var range in rangePairs.Select(r => Range(r.Trim())).Where(range => range != null))
            {
                retVal.Add(range);
            }
            return retVal;
        }

        internal XLIdManager ShapeIdManager { get; private set; }

        // Used by Janitor.Fody
        private void DisposeManaged()
        {
            Worksheets.ForEach(w => (w as XLWorksheet).Cleanup());
        }


        public void Dispose()
        {
            // Leave this empty so that Janitor.Fody can do its work
        }


        public Boolean Use1904DateSystem { get; set; }

        public XLWorkbook SetUse1904DateSystem()
        {
            return SetUse1904DateSystem(true);
        }

        public XLWorkbook SetUse1904DateSystem(Boolean value)
        {
            Use1904DateSystem = value;
            return this;
        }

        public IXLWorksheet AddWorksheet()
        {
            return Worksheets.Add();
        }

        public IXLWorksheet AddWorksheet(Int32 position)
        {
            return Worksheets.Add(position);
        }

        public IXLWorksheet AddWorksheet(String sheetName)
        {
            return Worksheets.Add(sheetName);
        }

        public IXLWorksheet AddWorksheet(String sheetName, Int32 position)
        {
            return Worksheets.Add(sheetName, position);
        }

        public void AddWorksheet(DataSet dataSet)
        {
            Worksheets.Add(dataSet);
        }

        public void AddWorksheet(IXLWorksheet worksheet)
        {
            worksheet.CopyTo(this, worksheet.Name);
        }

        public IXLWorksheet AddWorksheet(DataTable dataTable)
        {
            return Worksheets.Add(dataTable);
        }

        public IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName)
        {
            return Worksheets.Add(dataTable, sheetName);
        }

        public IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName, String tableName)
        {
            return Worksheets.Add(dataTable, sheetName, tableName);
        }

        private XLCalcEngine _calcEngine;

        internal XLCalcEngine CalcEngine
        {
            get { return _calcEngine ??= new XLCalcEngine(CultureInfo.CurrentCulture); }
        }

        public XLCellValue Evaluate(String expression)
        {
            return CalcEngine.EvaluateFormula(expression, this).ToCellValue();
        }

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        public void RecalculateAllFormulas()
        {
            foreach (var sheet in WorksheetsInternal)
                sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(XLSheetRange.Full);

            CalcEngine.Recalculate(this, null);
        }

        private static XLCalcEngine _calcEngineExpr;
        private SpreadsheetDocumentType _spreadsheetDocumentType;

        private static XLCalcEngine CalcEngineExpr
        {
            get { return _calcEngineExpr ??= new XLCalcEngine(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Evaluate a formula and return a value. Formulas with references don't work and culture used for conversion is invariant.
        /// </summary>
        public static XLCellValue EvaluateExpr(String expression)
        {
            return CalcEngineExpr.EvaluateFormula(expression).ToCellValue();
        }

        public String Author { get; set; }

        public Boolean LockStructure
        {
            get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure);
            set
            {
                if (!Protection.IsProtected)
                    throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockStructure)} property");

                Protection.AllowElement(XLWorkbookProtectionElements.Structure, value);
            }
        }

        public XLWorkbook SetLockStructure(Boolean value)
        {
            LockStructure = value; return this;
        }

        public Boolean LockWindows
        {
            get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows);
            set
            {
                if (!Protection.IsProtected)
                    throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockWindows)} property");

                Protection.AllowElement(XLWorkbookProtectionElements.Windows, value);
            }
        }

        public XLWorkbook SetLockWindows(Boolean value)
        {
            LockWindows = value; return this;
        }

        public Boolean IsPasswordProtected => Protection.IsPasswordProtected;
        public Boolean IsProtected => Protection.IsProtected;

        IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protection
        {
            get => Protection;
            set => Protection = value as XLWorkbookProtection;
        }

        internal XLWorkbookProtection Protection
        {
            get => _workbookProtection;
            set
            {
                _workbookProtection = value.Clone().CastTo<XLWorkbookProtection>();
            }
        }

        public IXLWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
        {
            return Protection.Protect(algorithm);
        }

        public IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements)
            => Protection.Protect(allowedElements);

        public IXLWorkbookProtection Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
            => Protection.Protect(algorithm, allowedElements);

        public IXLWorkbookProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm)

        {
            return Protect(password, algorithm, XLWorkbookProtectionElements.Windows);
        }

        public IXLWorkbookProtection Protect(String password, Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        {
            return Protection.Protect(password, algorithm, allowedElements);
        }

        IXLElementProtection IXLProtectable.Protect(Algorithm algorithm)
        {
            return Protect(algorithm);
        }

        IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm)
        {
            return Protect(password, algorithm);
        }

        IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(XLWorkbookProtectionElements allowedElements)
            => Protect(allowedElements);

        IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
            => Protect(algorithm, allowedElements);

        IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(string password, Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
            => Protect(password, algorithm, allowedElements);

        public IXLWorkbookProtection Unprotect()
        {
            return Protection.Unprotect();
        }

        public IXLWorkbookProtection Unprotect(String password)
        {
            return Protection.Unprotect(password);
        }

        IXLElementProtection IXLProtectable.Unprotect()
        {
            return Unprotect();
        }

        IXLElementProtection IXLProtectable.Unprotect(String password)
        {
            return Unprotect(password);
        }

        /// <summary>
        /// Notify various component of a workbook that sheet has been added.
        /// </summary>
        internal void NotifyWorksheetAdded(XLWorksheet newSheet)
        {
            _calcEngine.OnAddedSheet(newSheet);
        }

        /// <summary>
        /// Notify various component of a workbook that sheet is about to be removed.
        /// </summary>
        internal void NotifyWorksheetDeleting(XLWorksheet sheet)
        {
            _calcEngine.OnDeletingSheet(sheet);
        }

        public override string ToString()
        {
            switch (_loadSource)
            {
                case XLLoadSource.New:
                    return "XLWorkbook(new)";

                case XLLoadSource.File:
                    return String.Format("XLWorkbook({0})", _originalFile);

                case XLLoadSource.Stream:
                    return String.Format("XLWorkbook({0})", _originalStream.ToString());

                default:
                    throw new NotImplementedException();
            }
        }
    }
}
