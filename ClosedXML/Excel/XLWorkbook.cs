using ClosedXML.Excel.CalcEngine;
using ClosedXML.Extensions;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

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
            get { return WorksheetsInternal; }
        }

        /// <summary>
        ///   Gets an object to manipulate this workbook's named ranges.
        /// </summary>
        public IXLNamedRanges NamedRanges { get; private set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        public IXLTheme Theme { get; private set; }

        public IXLPivotSources PivotSources { get; private set; }

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

        internal XLColor GetXLColor(XLThemeColor themeColor)
        {
            switch (themeColor)
            {
                case XLThemeColor.Text1:
                    return Theme.Text1;

                case XLThemeColor.Background1:
                    return Theme.Background1;

                case XLThemeColor.Text2:
                    return Theme.Text2;

                case XLThemeColor.Background2:
                    return Theme.Background2;

                case XLThemeColor.Accent1:
                    return Theme.Accent1;

                case XLThemeColor.Accent2:
                    return Theme.Accent2;

                case XLThemeColor.Accent3:
                    return Theme.Accent3;

                case XLThemeColor.Accent4:
                    return Theme.Accent4;

                case XLThemeColor.Accent5:
                    return Theme.Accent5;

                case XLThemeColor.Accent6:
                    return Theme.Accent6;

                default:
                    throw new ArgumentException("Invalid theme color");
            }
        }

        public IXLNamedRange NamedRange(String rangeName)
        {
            if (rangeName.Contains("!"))
            {
                var split = rangeName.Split('!');
                var first = split[0];
                var wsName = first.StartsWith("'") ? first.Substring(1, first.Length - 2) : first;
                var name = split[1];
                if (TryGetWorksheet(wsName, out IXLWorksheet ws))
                {
                    var range = ws.NamedRange(name);
                    return range ?? NamedRange(name);
                }
                return null;
            }
            return NamedRanges.NamedRange(rangeName);
        }

        public Boolean TryGetWorksheet(String name, out IXLWorksheet worksheet)
        {
            return Worksheets.TryGetWorksheet(name, out worksheet);
        }

        public IXLRange RangeFromFullAddress(String rangeAddress, out IXLWorksheet ws)
        {
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

        public IXLCell CellFromFullAddress(String cellAddress, out IXLWorksheet ws)
        {
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
            PathHelper.CreateDirectory(Path.GetDirectoryName(file));
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
                        // not really nessesary, because I changed CopyStream too.
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
            var table = this.Worksheets
                .SelectMany(ws => ws.Tables)
                .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

            if (table == null)
                throw new ArgumentOutOfRangeException($"Table {tableName} was not found.");

            return table;
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
        /// <returns></returns>
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

        #endregion Fields

        #region Constructor

        /// <summary>
        ///   Creates a new Excel workbook.
        /// </summary>
        public XLWorkbook()
            : this(XLEventTracking.Enabled)
        {
        }

        internal XLWorkbook(String file, Boolean asTemplate)
            : this(XLEventTracking.Enabled)
        {
            LoadSheetsFromTemplate(file);
        }

        public XLWorkbook(XLEventTracking eventTracking)
        {
            EventTracking = eventTracking;
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
            PivotSources = new XLPivotSources();
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
        public XLWorkbook(String file)
            : this(file, XLEventTracking.Enabled)
        {
        }

        public XLWorkbook(String file, XLEventTracking eventTracking)
            : this(eventTracking)
        {
            _loadSource = XLLoadSource.File;
            _originalFile = file;
            _spreadsheetDocumentType = GetSpreadsheetDocumentType(_originalFile);
            Load(file);
        }

        public XLWorkbook(String file, LoadOptions loadOptions)
            : this(file, loadOptions.EventTracking)
        {
            if (loadOptions.RecalculateAllFormulas)
                this.RecalculateAllFormulas();
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
                this.RecalculateAllFormulas();
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
            var namedRange = NamedRange(namedCell);
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
            var namedRange = NamedRange(range);
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

        public void Dispose()
        {
            Worksheets.ForEach(w => (w as XLWorksheet).Cleanup());
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

        public IXLWorksheet AddWorksheet(DataTable dataTable)
        {
            return Worksheets.Add(dataTable);
        }

        public void AddWorksheet(DataSet dataSet)
        {
            Worksheets.Add(dataSet);
        }

        public void AddWorksheet(IXLWorksheet worksheet)
        {
            worksheet.CopyTo(this, worksheet.Name);
        }

        public IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName)
        {
            return Worksheets.Add(dataTable, sheetName);
        }

        private XLCalcEngine _calcEngine;

        private XLCalcEngine CalcEngine
        {
            get { return _calcEngine ?? (_calcEngine = new XLCalcEngine(this)); }
        }

        public Object Evaluate(String expression)
        {
            return CalcEngine.Evaluate(expression);
        }

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        public void RecalculateAllFormulas()
        {
            InvalidateFormulas();
            Worksheets.ForEach(sheet => sheet.RecalculateAllFormulas());
        }

        private static XLCalcEngine _calcEngineExpr;
        private SpreadsheetDocumentType _spreadsheetDocumentType;

        private static XLCalcEngine CalcEngineExpr
        {
            get { return _calcEngineExpr ?? (_calcEngineExpr = new XLCalcEngine()); }
        }

        public static Object EvaluateExpr(String expression)
        {
            return CalcEngineExpr.Evaluate(expression);
        }

        public String Author { get; set; }

        public Boolean LockStructure { get; set; }

        public XLWorkbook SetLockStructure(Boolean value)
        {
            LockStructure = value; return this;
        }

        public Boolean LockWindows { get; set; }

        public XLWorkbook SetLockWindows(Boolean value)
        {
            LockWindows = value; return this;
        }

        internal HexBinaryValue LockPassword { get; set; }
        public Boolean IsPasswordProtected { get { return LockPassword != null; } }

        public void Protect(Boolean lockStructure, Boolean lockWindows, String workbookPassword)
        {
            if (IsPasswordProtected && workbookPassword == null)
                throw new InvalidOperationException("The workbook is password protected");

            var hashPassword = workbookPassword.HashPassword();
            if (IsPasswordProtected && LockPassword != hashPassword)
                throw new ArgumentException("Invalid password");

            if (IsPasswordProtected && (lockStructure || lockWindows))
                throw new InvalidOperationException("The workbook is already protected");

            if (IsPasswordProtected && hashPassword != null && !lockStructure && !lockWindows)
            {
                // Workbook currently protected, but we're unsetting the 2 flags
                // Hence unprotect workbook using password.
                LockPassword = null;
            }

            if (!IsPasswordProtected && hashPassword != null && (lockStructure || lockWindows))
            {
                //Protect workbook using password.
                LockPassword = hashPassword;
            }

            LockStructure = lockStructure;
            LockWindows = lockWindows;
        }

        public void Protect()
        {
            Protect(true);
        }

        public void Protect(string workbookPassword)
        {
            Protect(true, false, workbookPassword);
        }

        public void Protect(Boolean lockStructure)
        {
            Protect(lockStructure, false);
        }

        public void Protect(Boolean lockStructure, Boolean lockWindows)
        {
            Protect(lockStructure, lockWindows, null);
        }

        public void Unprotect()
        {
            Protect(false, false);
        }

        public void Unprotect(string workbookPassword)
        {
            Protect(false, false, workbookPassword);
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

        public void SuspendEvents()
        {
            foreach (var ws in WorksheetsInternal)
            {
                ws.SuspendEvents();
            }
        }

        public void ResumeEvents()
        {
            foreach (var ws in WorksheetsInternal)
            {
                ws.ResumeEvents();
            }
        }
    }
}
