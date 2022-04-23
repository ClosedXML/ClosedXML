using ClosedXML.Excel.CalcEngine;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    public enum XLEventTracking { Enabled, Disabled }

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

        public static double DefaultRowHeight { get; private set; }
        public static double DefaultColumnWidth { get; private set; }

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

        public static XLWorkbook OpenFromTemplate(string path)
        {
            return new XLWorkbook(path, asTemplate: true);
        }

        #endregion Static

        private bool _disposed = false;

        internal readonly List<UnsupportedSheet> UnsupportedSheets =
            new List<UnsupportedSheet>();

        public XLEventTracking EventTracking { get; set; }

        /// <summary>
        /// Counter increasing at workbook data change. Serves to determine if the cell formula
        /// has to be recalculated.
        /// </summary>
        internal long RecalculationCounter { get; private set; }

        /// <summary>
        /// Notify that workbook data has been changed which means that cached formula values
        /// need to be re-evaluated.
        /// </summary>
        internal void InvalidateFormulas()
        {
            RecalculationCounter++;
        }

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
            get
            {
                ThrowIfDisposed();

                return WorksheetsInternal;
            }
        }

        /// <summary>
        ///   Gets an object to manipulate this workbook's named ranges.
        /// </summary>
        public IXLNamedRanges NamedRanges { get; private set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        public IXLTheme Theme { get; private set; }

        /// <summary>
        ///   Gets or sets the default style for the workbook.
        ///   <para>All new worksheets will use this style.</para>
        /// </summary>
        public IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets or sets the default row height for the workbook.
        ///   <para>All new worksheets will use this row height.</para>
        /// </summary>
        public double RowHeight { get; set; }

        /// <summary>
        ///   Gets or sets the default column width for the workbook.
        ///   <para>All new worksheets will use this column width.</para>
        /// </summary>
        public double ColumnWidth { get; set; }

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

        public bool CalculationOnSave { get; set; }
        public bool ForceFullCalculation { get; set; }
        public bool FullCalculationOnLoad { get; set; }
        public bool FullPrecision { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's reference style.
        /// </summary>
        public XLReferenceStyle ReferenceStyle { get; set; }

        public IXLCustomProperties CustomProperties { get; private set; }

        public bool ShowFormulas { get; set; }
        public bool ShowGridLines { get; set; }
        public bool ShowOutlineSymbols { get; set; }
        public bool ShowRowColHeaders { get; set; }
        public bool ShowRuler { get; set; }
        public bool ShowWhiteSpace { get; set; }
        public bool ShowZeros { get; set; }
        public bool RightToLeft { get; set; }

        public bool DefaultShowFormulas
        {
            get { return false; }
        }

        public bool DefaultShowGridLines
        {
            get { return true; }
        }

        public bool DefaultShowOutlineSymbols
        {
            get { return true; }
        }

        public bool DefaultShowRowColHeaders
        {
            get { return true; }
        }

        public bool DefaultShowRuler
        {
            get { return true; }
        }

        public bool DefaultShowWhiteSpace
        {
            get { return true; }
        }

        public bool DefaultShowZeros
        {
            get { return true; }
        }

        public IXLFileSharing FileSharing { get; } = new XLFileSharing();

        public bool DefaultRightToLeft
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

        public IXLNamedRange NamedRange(string rangeName)
        {
            ThrowIfDisposed();

            if (rangeName.Contains("!"))
            {
                var split = rangeName.Split('!');
                var first = split[0];
                var wsName = first.StartsWith("'") ? first.Substring(1, first.Length - 2) : first;
                var name = split[1];
                if (TryGetWorksheet(wsName, out var ws))
                {
                    var range = ws.NamedRange(name);
                    return range ?? NamedRange(name);
                }
                return null;
            }
            return NamedRanges.NamedRange(rangeName);
        }

        public bool TryGetWorksheet(string name, out IXLWorksheet worksheet)
        {
            ThrowIfDisposed();

            return Worksheets.TryGetWorksheet(name, out worksheet);
        }

        public IXLRange RangeFromFullAddress(string rangeAddress, out IXLWorksheet ws)
        {
            ThrowIfDisposed();

            ws = null;
            if (!rangeAddress.Contains('!')) return null;

            var split = rangeAddress.Split('!');
            var wsName = split[0].UnescapeSheetName();
            if (TryGetWorksheet(wsName, out ws))
            {
                return ws.Range(split[1]);
            }
            return null;
        }

        public IXLCell CellFromFullAddress(string cellAddress, out IXLWorksheet ws)
        {
            ThrowIfDisposed();

            ws = null;
            if (!cellAddress.Contains('!')) return null;

            var split = cellAddress.Split('!');
            var wsName = split[0].UnescapeSheetName();
            if (TryGetWorksheet(wsName, out ws))
            {
                return ws.Cell(split[1]);
            }
            return null;
        }

        /// <summary>
        ///   Saves the current workbook.
        /// </summary>
        public void Save()
        {
            ThrowIfDisposed();

#if DEBUG
            Save(true, false);
#else
            Save(false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook and optionally performs validation
        /// </summary>
        public void Save(bool validate, bool evaluateFormulae = false)
        {
            ThrowIfDisposed();

            Save(new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void Save(SaveOptions options)
        {
            ThrowIfDisposed();

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
        public void SaveAs(string file)
        {
            ThrowIfDisposed();

#if DEBUG
            SaveAs(file, true, false);
#else
            SaveAs(file, false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook to a file and optionally validates it.
        /// </summary>
        public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
        {
            ThrowIfDisposed();

            SaveAs(file, new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void SaveAs(string file, SaveOptions options)
        {
            ThrowIfDisposed();

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
                if (string.Compare(_originalFile.Trim(), file.Trim(), true) != 0)
                {
                    File.Copy(_originalFile, file, true);
                    File.SetAttributes(file, FileAttributes.Normal);
                }

                CreatePackage(file, GetSpreadsheetDocumentType(file), options);
            }
            else if (_loadSource == XLLoadSource.Stream)
            {
                _originalStream.Position = 0;

                using var fileStream = File.Create(file);
                CopyStream(_originalStream, fileStream);
                CreatePackage(fileStream, false, _spreadsheetDocumentType, options);
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
                    throw new ArgumentException(string.Format("Extension '{0}' is not supported. Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.", extension));
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
            ThrowIfDisposed();

#if DEBUG
            SaveAs(stream, true, false);
#else
            SaveAs(stream, false, false);
#endif
        }

        /// <summary>
        ///   Saves the current workbook to a stream and optionally validates it.
        /// </summary>
        public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
        {
            ThrowIfDisposed();

            SaveAs(stream, new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true
            });
        }

        public void SaveAs(Stream stream, SaveOptions options)
        {
            ThrowIfDisposed();

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
                    using var ms = new MemoryStream();
                    CreatePackage(ms, true, _spreadsheetDocumentType, options);
                    // not really nessesary, because I changed CopyStream too.
                    // but for better understanding and if somebody in the future
                    // provide an changed version of CopyStream
                    ms.Position = 0;
                    CopyStream(ms, stream);
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
            ThrowIfDisposed();

            var table = Worksheets
                .SelectMany(ws => ws.Tables)
                .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

            if (table == null)
                throw new ArgumentOutOfRangeException($"Table {tableName} was not found.");

            return table;
        }

        public IXLWorksheet Worksheet(string name)
        {
            ThrowIfDisposed();

            return WorksheetsInternal.Worksheet(name);
        }

        public IXLWorksheet Worksheet(int position)
        {
            ThrowIfDisposed();

            return WorksheetsInternal.Worksheet(position);
        }

        public IXLCustomProperty CustomProperty(string name)
        {
            ThrowIfDisposed();

            return CustomProperties.CustomProperty(name);
        }

        public IXLCells FindCells(Func<IXLCell, bool> predicate)
        {
            ThrowIfDisposed();

            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (var ws in WorksheetsInternal)
            {
                foreach (XLCell cell in ws.CellsUsed(XLCellsUsedOptions.All))
                {
                    if (predicate(cell))
                        cells.Add(cell);
                }
            }
            return cells;
        }

        public IXLRows FindRows(Func<IXLRow, bool> predicate)
        {
            ThrowIfDisposed();

            var rows = new XLRows(worksheet: null);
            foreach (var ws in WorksheetsInternal)
            {
                foreach (var row in ws.Rows().Where(predicate))
                    rows.Add(row as XLRow);
            }
            return rows;
        }

        public IXLColumns FindColumns(Func<IXLColumn, bool> predicate)
        {
            ThrowIfDisposed();

            var columns = new XLColumns(worksheet: null);
            foreach (var ws in WorksheetsInternal)
            {
                foreach (var column in ws.Columns().Where(predicate))
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
        /// <returns></returns>
        public IEnumerable<IXLCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false)
        {
            ThrowIfDisposed();

            foreach (var ws in WorksheetsInternal)
            {
                foreach (var cell in ws.Search(searchText, compareOptions, searchFormulae))
                    yield return cell;
            }
        }

        #region Fields

        private XLLoadSource _loadSource = XLLoadSource.New;
        private string _originalFile;
        private Stream _originalStream;
        private XLWorkbookProtection _workbookProtection;

        #endregion Fields

        #region Constructor

        /// <summary>
        ///   Creates a new Excel workbook.
        /// </summary>
        public XLWorkbook()
            : this(XLEventTracking.Enabled)
        {
        }

        internal XLWorkbook(string file, bool asTemplate)
            : this(XLEventTracking.Enabled)
        {
            LoadSheetsFromTemplate(file);
        }

        public XLWorkbook(XLEventTracking eventTracking)
        {
            EventTracking = eventTracking;
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
            NamedRanges = new XLNamedRanges(this);
            CustomProperties = new XLCustomProperties(this);
            ShapeIdManager = new XLIdManager();
            Author = Environment.UserName;
        }

        public XLWorkbook(LoadOptions loadOptions)
            : this(loadOptions.EventTracking)
        {
        }

        /// <summary>
        ///   Opens an existing workbook from a file.
        /// </summary>
        /// <param name = "file">The file to open.</param>
        public XLWorkbook(string file)
            : this(file, XLEventTracking.Enabled)
        {
        }

        public XLWorkbook(string file, XLEventTracking eventTracking)
            : this(eventTracking)
        {
            _loadSource = XLLoadSource.File;
            _originalFile = file;
            _spreadsheetDocumentType = GetSpreadsheetDocumentType(_originalFile);
            Load(file);
        }

        public XLWorkbook(string file, LoadOptions loadOptions)
            : this(file, loadOptions.EventTracking)
        {
            if (loadOptions.RecalculateAllFormulas)
                RecalculateAllFormulas();
        }

        /// <summary>
        ///   Opens an existing workbook from a stream.
        /// </summary>
        /// <param name = "stream">The stream to open.</param>
        public XLWorkbook(Stream stream)
            : this(stream, XLEventTracking.Enabled)
        {
        }

        public XLWorkbook(Stream stream, XLEventTracking eventTracking)
            : this(eventTracking)
        {
            _loadSource = XLLoadSource.Stream;
            _originalStream = stream;
            Load(stream);
        }

        public XLWorkbook(Stream stream, LoadOptions loadOptions)
            : this(stream, loadOptions.EventTracking)
        {
            if (loadOptions.RecalculateAllFormulas)
                RecalculateAllFormulas();
        }

        #endregion Constructor

        #region Nested type: UnsupportedSheet

        internal sealed class UnsupportedSheet
        {
            public bool IsActive;
            public uint SheetId;
            public int Position;
        }

        #endregion Nested type: UnsupportedSheet

        public IXLCell Cell(string namedCell)
        {
            ThrowIfDisposed();

            var namedRange = NamedRange(namedCell);
            if (namedRange != null)
            {
                return namedRange.Ranges?.FirstOrDefault()?.FirstCell();
            }
            else
                return CellFromFullAddress(namedCell, out _);
        }

        public IXLCells Cells(string namedCells)
        {
            ThrowIfDisposed();

            return Ranges(namedCells).Cells();
        }

        public IXLRange Range(string range)
        {
            ThrowIfDisposed();

            var namedRange = NamedRange(range);
            if (namedRange != null)
                return namedRange.Ranges.FirstOrDefault();
            else
                return RangeFromFullAddress(range, out _);
        }

        public IXLRanges Ranges(string ranges)
        {
            ThrowIfDisposed();

            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (var range in rangePairs.Select(r => Range(r.Trim())).Where(range => range != null))
            {
                retVal.Add(range);
            }
            return retVal;
        }

        internal XLIdManager ShapeIdManager { get; private set; }

        public bool Use1904DateSystem { get; set; }

        public XLWorkbook SetUse1904DateSystem()
        {
            ThrowIfDisposed();

            return SetUse1904DateSystem(true);
        }

        public XLWorkbook SetUse1904DateSystem(bool value)
        {
            ThrowIfDisposed();

            Use1904DateSystem = value;
            return this;
        }

        public IXLWorksheet AddWorksheet()
        {
            ThrowIfDisposed();

            return Worksheets.Add();
        }

        public IXLWorksheet AddWorksheet(int position)
        {
            ThrowIfDisposed();

            return Worksheets.Add(position);
        }

        public IXLWorksheet AddWorksheet(string sheetName)
        {
            ThrowIfDisposed();

            return Worksheets.Add(sheetName);
        }

        public IXLWorksheet AddWorksheet(string sheetName, int position)
        {
            ThrowIfDisposed();

            return Worksheets.Add(sheetName, position);

        }

        public IXLWorksheet AddWorksheet(DataTable dataTable)
        {
            ThrowIfDisposed();

            return Worksheets.Add(dataTable);
        }

        public void AddWorksheet(DataSet dataSet)
        {
            ThrowIfDisposed();

            Worksheets.Add(dataSet);
        }

        public void AddWorksheet(IXLWorksheet worksheet)
        {
            ThrowIfDisposed();

            worksheet.CopyTo(this, worksheet.Name);
        }

        public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName)
        {
            ThrowIfDisposed();

            return Worksheets.Add(dataTable, sheetName);
        }

        private XLCalcEngine _calcEngine;

        private XLCalcEngine CalcEngine
        {
            get { return _calcEngine ?? (_calcEngine = new XLCalcEngine(this)); }
        }

        public object Evaluate(string expression)
        {
            ThrowIfDisposed();

            return CalcEngine.Evaluate(expression);
        }

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        public void RecalculateAllFormulas()
        {
            ThrowIfDisposed();

            InvalidateFormulas();
            Worksheets.ForEach(sheet => sheet.RecalculateAllFormulas());
        }

        private SpreadsheetDocumentType _spreadsheetDocumentType;

        private static XLCalcEngine _calcEngineExpr;

        private static XLCalcEngine CalcEngineExpr
        {
            get { return _calcEngineExpr ?? (_calcEngineExpr = new XLCalcEngine()); }
        }

        public static object EvaluateExpr(string expression)
        {
            return CalcEngineExpr.Evaluate(expression);
        }

        public string Author { get; set; }

        public bool LockStructure
        {
            get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure);
            set
            {
                ThrowIfDisposed();

                if (!Protection.IsProtected)
                    throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockStructure)} property");

                Protection.AllowElement(XLWorkbookProtectionElements.Structure, value);
            }
        }

        public XLWorkbook SetLockStructure(bool value)
        {
            ThrowIfDisposed();

            LockStructure = value; return this;
        }

        public bool LockWindows
        {
            get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows);
            set
            {
                ThrowIfDisposed();

                if (!Protection.IsProtected)
                    throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockWindows)} property");

                Protection.AllowElement(XLWorkbookProtectionElements.Windows, value);
            }
        }

        public XLWorkbook SetLockWindows(bool value)
        {
            ThrowIfDisposed();

            LockWindows = value; return this;
        }

        public bool IsPasswordProtected => Protection.IsPasswordProtected;
        public bool IsProtected => Protection.IsProtected;

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
                ThrowIfDisposed();

                _workbookProtection = value.Clone().CastTo<XLWorkbookProtection>();
            }
        }

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        public IXLWorkbookProtection Protect(bool lockStructure, bool lockWindows, string password)
        {
            ThrowIfDisposed();

            var allowedElements = XLWorkbookProtectionElements.Everything;

            var protection = Protection.Protect(password, DefaultProtectionAlgorithm, allowedElements);

            if (lockStructure)
                protection.DisallowElement(XLWorkbookProtectionElements.Structure);

            if (lockWindows)
                protection.DisallowElement(XLWorkbookProtectionElements.Windows);

            return protection;
        }

        public IXLWorkbookProtection Protect()
        {
            ThrowIfDisposed();

            return Protection.Protect();
        }

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        public IXLWorkbookProtection Protect(bool lockStructure)
        {
            ThrowIfDisposed();

            return Protect(lockStructure, lockWindows: false, password: null);
        }

        [Obsolete("Use Protect(String password, Algorithm algorithm, TElement allowedElements)")]
        public IXLWorkbookProtection Protect(bool lockStructure, bool lockWindows)
        {
            ThrowIfDisposed();

            return Protect(lockStructure, lockWindows, null);
        }

        public IXLWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm)
        {
            ThrowIfDisposed();

            return Protect(password, algorithm, XLWorkbookProtectionElements.Windows);
        }

        public IXLWorkbookProtection Protect(string password, Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        {
            ThrowIfDisposed();

            return Protection.Protect(password, algorithm, allowedElements);
        }

        IXLElementProtection IXLProtectable.Protect()
        {
            ThrowIfDisposed();

            return Protect();
        }

        IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm)
        {
            ThrowIfDisposed();

            return Protect(password, algorithm);
        }

        public IXLWorkbookProtection Unprotect()
        {
            ThrowIfDisposed();

            return Protection.Unprotect();
        }

        public IXLWorkbookProtection Unprotect(string password)
        {
            ThrowIfDisposed();

            return Protection.Unprotect(password);
        }

        IXLElementProtection IXLProtectable.Unprotect()
        {
            ThrowIfDisposed();

            return Unprotect();
        }

        IXLElementProtection IXLProtectable.Unprotect(string password)
        {
            ThrowIfDisposed();

            return Unprotect(password);
        }

        public override string ToString()
        {
            switch (_loadSource)
            {
                case XLLoadSource.New:
                    return "XLWorkbook(new)";

                case XLLoadSource.File:
                    return string.Format("XLWorkbook({0})", _originalFile);

                case XLLoadSource.Stream:
                    return string.Format("XLWorkbook({0})", _originalStream.ToString());

                default:
                    throw new NotImplementedException();
            }
        }

        public void SuspendEvents()
        {
            ThrowIfDisposed();

            foreach (var ws in WorksheetsInternal)
            {
                ws.SuspendEvents();
            }
        }

        public void ResumeEvents()
        {
            ThrowIfDisposed();

            foreach (var ws in WorksheetsInternal)
            {
                ws.ResumeEvents();
            }
        }

        public void Dispose()
        {
            // Dispose of unmanaged resources.
            Dispose(true);
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                Worksheets.ForEach(w => (w as XLWorksheet).Cleanup());
            }

            _disposed = true;
        }

        void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException("TemplateClass");
            }
        }
    }
}
